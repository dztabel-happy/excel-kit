"""Run in an empty output directory. Python orchestrates the public ExcelKit CLI."""
import argparse
import json
from pathlib import Path
import subprocess


def main(directory, cli):
    directory.mkdir(parents=True, exist_ok=True)
    source, baseline, working, delivery = [str(directory / (name + ".xlsx"))
        for name in ("source", "baseline", "working", "delivery")]

    def invoke(command, request):
        p = subprocess.run([*cli, command, "-"], input=json.dumps(request), text=True, capture_output=True)
        result = json.loads(p.stdout)
        if not result["ok"]:
            raise RuntimeError(result)
        return result

    def write(sheet, cell, *rows):
        return dict(op="write", sheet=sheet, range=cell, values=rows)

    def formula(text):
        return {"formula": "=" + text}

    def name(sheet, address):
        return dict(op="name.set", name="Rate", reference=sheet + "!" + address)

    chart = dict(op="chart.create", sheet="Model", name="Charges", type="column",
        series=[dict(categories="Model!A2:A3", values="Model!C2:C3")],
        anchor=dict(kind="oneCell", **{"from": {"cell": "F2"}}, width=260, height=140))
    original = invoke("run", dict(create=dict(sheets=["Model"]), output=source, operations=[
        write("Model", "A1", ["Review\ncharges\nby region"]),
        dict(op="merge", sheet="Model", range="A1:C1"),
        dict(op="format", sheet="Model", range="A1", style=dict(alignment=dict(wrapText=True),font=dict(bold=True))),
        write("Model", "A2", [10,2,formula("A2*B2*Rate")],[7,3,formula("$A2+A$2+$A$2")]),
        write("Model", "H1", [4]), name("Model","$H$1"),chart]))
    invoke("run",dict(create=dict(sheets=["Model","Data","Summary"]), output=baseline, operations=[
        write("Model","A2",[5,None,4]),write("Model","A6",[3]), write("Model","B1",[1.5]),name("Model","$B$1"),
        write("Model","E6",[formula("C6*D6*Rate")]), write("Model","C6",[10,2]),
        dict(op="layout",sheet="Model",rows=[dict(start=5,count=1,height=15)]),
        write("Data","A1",["Region","Cost","Qty"],["East",-1.25,1],["East",-1.20,1],["West",3,0]),
        dict(op="table.create",sheet="Data",name="Inputs",range="A1:C4")]))
    broken = invoke("run",dict(input=baseline,output=working,operations=[write("Model","E6",[30])]))
    difference = invoke("compare",dict(left=baseline,right=working,sheet="Model",range="C5:E7"))
    assert difference["total"] == 1 and difference["items"][0]["left"]["content"]["type"] == "formula", difference

    # A receiving session can contain other work before the repair starts.
    with subprocess.Popen([*cli,"session"],stdin=subprocess.PIPE,stdout=subprocess.PIPE,text=True) as proc:
        seq = 0
        def rpc(method, **params):
            nonlocal seq
            seq += 1
            message = dict(jsonrpc="2.0",id=seq,method="session."+method,params=dict(sessionId="model",**params))
            proc.stdin.write(json.dumps(message) + "\n")
            proc.stdin.flush()
            response = json.loads(proc.stdout.readline())["result"]["response"]
            if not response["ok"]:
                raise RuntimeError(response)
            return response

        rpc("open",input=working,output=delivery,expectedInputVersion=broken["output"]["version"],reread="precise")
        edit = rpc("run",**write("Summary","A9",["Reviewed"]))
        version = edit["outputVersion"]
        move = dict(op="range.transfer",sheet="Model",cell="C5",chart="Charges",
            source=dict(path=source,sheet="Model",range="A1:C3",expectedVersion=original["output"]["version"]),
            bindings=[dict(**{"from":"Rate","to":"Rate"})],replaceRange="C5:E7",expectedTargetVersion=version)
        move_script = "const step = await wb.run(args.move); return step.result.rangeOperations[0].transfer;"
        # The exact replacement can be tried against the current unsaved model.
        plan = rpc("script",expectedSessionVersion=version,trial=True,strict=True,script=move_script,arguments={"move":move})
        assert plan["outputVersion"] == version
        rpc("run",expectedSessionVersion=version,**move)
        scenario = """
const step = await wb.run({operations:[{op:"write",sheet:"Model",range:"C6",values:[[args.units]]}],
  calculation:{mode:"targets",targets:[{sheet:"Model",range:"E6"}]},read:[{sheet:"Model",range:"E6"}]});
return step.result.ranges[0].cells[0][0].result.value;
"""
        values = [rpc("script",trial=True,strict=True,script=scenario,arguments=dict(units=n))["result"] for n in (12,18)]
        assert values == [36,54], values
        chosen = rpc("script",strict=True,script=scenario,arguments=dict(units=12))
        assert chosen["result"] == 36
        rpc("run",op="pivot.create",sheet="Summary",cell="A1",name="Costs",source="Inputs",rows=["Region"],
            calculatedFields=[dict(name="UnitCost",formula="IF(Qty=0,0,ROUND(ABS(Cost/Qty),2))")],
            values=[dict(field="UnitCost",name="Unit cost")])
        rpc("run",operations=[write("Data","B2",[-1.27]),dict(op="pivot.refresh",sheet="Summary",name="Costs"),
            dict(op="chart.refresh",sheet="Model",name="Charges")])
        layout = dict(sheet="Model",range="C5:E5",axis="textClipping")
        clipped = rpc("query",layout=[layout])["layout"][0]["items"]
        assert len(clipped) == 1 and clipped[0]["range"] == "C5:E5", clipped
        rpc("run",op="layout",sheet="Model",rows=[dict(start=5,count=1,height=70)])
        assert rpc("query",layout=[layout])["layout"][0]["total"] == 0
        final = rpc("read",ranges=[dict(sheet="Model",range="E6:E7"),dict(sheet="Summary",range="B2:B4")],calculation=dict(mode="all"))
        values_seen = [[r[0]["result"]["value"] if r[0].get("result") else r[0]["input"]["value"] for r in area["cells"]] for area in final["ranges"]]
        assert values_seen == [[36,12],[1.24,0,0.27]], values_seen
        receipt = rpc("save")
        rpc("close")
        proc.stdin.close()
    report = dict(changedCell=difference["items"][0],transfer=plan["result"],scenarios=values,selectedUnits=12,
        committedCharge=chosen["result"],reread=chosen["reread"],clipping=clipped,pivotAndFormulaValues=values_seen,output=receipt["output"])
    (directory/"report.json").write_text(json.dumps(report,indent=2)+"\n")


if __name__ == "__main__":
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory",type=Path)
    parser.add_argument("--cli",nargs="+",default=["excel-kit"])
    args=parser.parse_args()
    main(args.directory.resolve(),args.cli)
