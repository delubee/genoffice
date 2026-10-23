"""Visible standard-library analysis for GenOffice typed workbook snapshots.

Run this from the native workspace terminal after workbook_export. It never
opens or writes an XLSX, installs dependencies, or calls an AI service.
"""

import argparse
from decimal import Decimal
import hashlib
import json
import math
import os
from pathlib import Path
import re
import uuid

MAX_BYTES = 64 * 1024 * 1024
MAX_RECORD = 256 * 1024


def read_manifest(path):
    encoded = path.read_bytes()
    if len(encoded) > MAX_RECORD:
        raise ValueError("manifest exceeds 256 KiB")
    manifest = json.loads(encoded.decode("utf-8"))
    if manifest.get("schemaVersion") != 1 or manifest.get("kind") != "workbook-snapshot" or manifest.get("complete") is not True:
        raise ValueError("a complete schemaVersion 1 snapshot is required")
    shape = manifest["shape"]
    if shape["rows"] * shape["columns"] != shape["cellCount"] or not 0 < shape["cellCount"] <= 1_000_000:
        raise ValueError("invalid source shape")
    name = manifest["data"]["fileName"]
    if not re.fullmatch(r"[a-zA-Z0-9_-]+\.ndjson", name):
        raise ValueError("source data must be a relative NDJSON file name")
    return manifest, hashlib.sha256(encoded).hexdigest()


def input_records(path, manifest):
    digest = hashlib.sha256()
    total = records = row = column = 0
    shape = manifest["shape"]
    with path.open("rb") as source:
        while True:
            line = source.readline(MAX_RECORD + 2)
            if not line:
                break
            total += len(line)
            if len(line) - 1 > MAX_RECORD or not line.endswith(b"\n") or total > MAX_BYTES:
                raise ValueError("source record or byte limit exceeded")
            digest.update(line)
            record = json.loads(line.decode("utf-8"), parse_constant=lambda _: (_ for _ in ()).throw(ValueError("nonfinite JSON value")))
            if record["rowOffset"] != row or record["columnOffset"] != column or not record["cells"]:
                raise ValueError("source records overlap or leave a gap")
            column += len(record["cells"])
            if row >= shape["rows"] or column > shape["columns"]:
                raise ValueError("source record exceeds its shape")
            if column == shape["columns"]:
                row += 1
                column = 0
            records += 1
            yield record
    expected = manifest["data"]
    if (row, column) != (shape["rows"], 0) or total != expected["bytes"] or records != expected["recordCount"] or digest.hexdigest() != expected["sha256"]:
        raise ValueError("source shape, digest, bytes or record count mismatch")


def summary_records(records, shape):
    current = []
    header = None
    totals = {}
    for record in records:
        current.extend(record["cells"])
        if len(current) != shape["columns"]:
            continue
        if header is None:
            header = [cell.get("value") for cell in current]
            if "Region" not in header or "Sales" not in header:
                raise ValueError("summary input requires Region and Sales headers")
        else:
            region_cell, sales_cell = current[header.index("Region")], current[header.index("Sales")]
            if region_cell["kind"] != "string":
                raise ValueError("Region must be a raw string")
            if sales_cell["kind"] == "formula":
                if sales_cell["calculationState"] != "ready" or sales_cell["cached"] is None:
                    raise ValueError("unverified formula cache cannot be used as a sales amount")
                sales_cell = sales_cell["cached"]
            if sales_cell["kind"] != "number":
                raise ValueError("Sales must be raw numeric data, not display text/date/error")
            region = region_cell["value"]
            totals[region] = totals.get(region, Decimal(0)) + Decimal(str(sales_cell["value"]))
        current = []
    if current or not totals:
        raise ValueError("source has no complete sales result")
    rows = [[{"kind": "string", "value": "Region"}, {"kind": "string", "value": "Sales"}]]
    order = [region for region in ["North", "South", "East", "West"] if region in totals]
    order.extend(sorted(set(totals) - set(order)))
    for region in order:
        amount = float(totals[region])
        if not math.isfinite(amount):
            raise ValueError("sales aggregate exceeds finite workbook precision")
        rows.append([{"kind": "string", "value": region}, {"kind": "number", "value": amount, "numberFormat": "#,##0.00"}])
    return ({"rowOffset": index, "columnOffset": 0, "cells": cells} for index, cells in enumerate(rows)), {"rows": len(rows), "columns": 2, "cellCount": len(rows) * 2}


def write_result(directory, records, shape, manifest, manifest_hash, sheet_name, allow_formulas):
    data_path = directory / "result.ndjson"
    temporary = directory / ("result." + str(uuid.uuid4()) + ".tmp")
    digest = hashlib.sha256()
    total = count = 0
    try:
        with temporary.open("xb") as target:
            for record in records:
                if not allow_formulas and any(cell["kind"] == "formula" for cell in record["cells"]):
                    raise ValueError("copying formulas requires explicit --allow-formulas")
                line = (json.dumps(record, ensure_ascii=False, allow_nan=False, separators=(",", ":")) + "\n").encode("utf-8")
                total += len(line)
                if len(line) - 1 > MAX_RECORD or total > MAX_BYTES:
                    raise ValueError("result exceeds record or 64 MiB data limit")
                target.write(line)
                digest.update(line)
                count += 1
            target.flush()
            os.fsync(target.fileno())
        temporary.replace(data_path)
        result = {"schemaVersion": 1, "kind": "workbook-result", "source": {
            "snapshotId": manifest["snapshotId"], "sourceManifestSha256": manifest_hash,
            "sourceDataSha256": manifest["data"]["sha256"], "sourceRevision": manifest["sourceRevision"]},
            "output": {"fileName": data_path.name, "sha256": digest.hexdigest(), "bytes": total, "recordCount": count, "encoding": "utf-8", "format": "typed-ndjson-v1"},
            "shape": shape, "target": {"newSheetName": sheet_name, "startCell": "A1"},
            "mode": "formulas" if allow_formulas else "values", "complete": True}
        destination = directory / "result.manifest.json"
        manifest_temporary = directory / ("manifest." + str(uuid.uuid4()) + ".tmp")
        with manifest_temporary.open("x", encoding="utf-8") as target:
            json.dump(result, target, ensure_ascii=False, allow_nan=False, separators=(",", ":"))
            target.flush()
            os.fsync(target.fileno())
        manifest_temporary.replace(destination)
        print(json.dumps({"resultManifestPath": str(destination), "shape": shape, "sourceRevision": manifest["sourceRevision"]}))
    finally:
        temporary.unlink(missing_ok=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_manifest", type=Path)
    parser.add_argument("--mode", choices=["summary", "copy"], default="summary")
    parser.add_argument("--new-sheet-name", default="Summary")
    parser.add_argument("--allow-formulas", action="store_true")
    args = parser.parse_args()
    path = args.input_manifest.resolve()
    manifest, manifest_hash = read_manifest(path)
    records = input_records(path.parent / manifest["data"]["fileName"], manifest)
    if args.mode == "summary":
        records, shape = summary_records(records, manifest["shape"])
    else:
        shape = manifest["shape"]
    write_result(path.parent, records, shape, manifest, manifest_hash, args.new_sheet_name, args.allow_formulas)


if __name__ == "__main__":
    main()
