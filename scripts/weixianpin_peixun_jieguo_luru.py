# -*- coding: utf-8 -*-
"""按「危险品培训结果录入」Skill：复制仓内一次性导入模板并写入当月名单。"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

REPO = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE = REPO / "assets" / "危险品录入" / "培训实施-一次性批量导入模版.xlsx"

ORG_NAME = "中国南方航空货运有限公司"
ORG_CREDIT = "91440101MA5D3DHK1J"
ORG_TYPE = "航空公司"
ID_TYPE = "个人专属编号"
DEFAULT_ATTEND = "正常"
DEFAULT_HOURS = 7


def _as_emp(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
    )


def _nation_code(v) -> str:
    if pd.isna(v):
        return ""
    t = str(v).strip()
    if t in {"中国", "CN", "CHN"}:
        return "CN"
    return t


def build_rows(
    zong_path: Path,
    caac_path: Path,
    shili_path: Path,
    train_type: str,
) -> pd.DataFrame:
    zong = pd.read_excel(zong_path)
    need = ["员工号", "姓名", "态度评价", "任务模拟", "考试", "评估得分", "评估结论"]
    miss = [c for c in need if c not in zong.columns]
    if miss:
        raise SystemExit(f"总成绩表缺列: {miss}")

    caac = pd.read_excel(caac_path)
    for c in ["姓名", "手机号码", "性别", "国籍"]:
        if c not in caac.columns:
            raise SystemExit(f"CAAC 导出缺列: {c}")
    caac = caac[["姓名", "手机号码", "性别", "国籍"]].drop_duplicates("姓名", keep="first")

    info = pd.read_excel(shili_path, sheet_name="人员信息")
    for c in ["员工号", "危险品"]:
        if c not in info.columns:
            raise SystemExit(f"实力表人员信息缺列: {c}")
    info = info.copy()
    info["员工号"] = _as_emp(info["员工号"])
    info = info[["员工号", "危险品"]].drop_duplicates("员工号", keep="first")

    df = zong[need].copy()
    df["员工号"] = _as_emp(df["员工号"])
    df = df.merge(caac, on="姓名", how="left")
    df = df.merge(info, on="员工号", how="left")

    if df["手机号码"].isna().any() or df["性别"].isna().any():
        bad = df[df["手机号码"].isna() | df["性别"].isna()][["员工号", "姓名"]]
        raise SystemExit(f"CAAC 未命中姓名:\n{bad.to_string(index=False)}")
    if df["危险品"].isna().any():
        bad = df[df["危险品"].isna()][["员工号", "姓名"]]
        raise SystemExit(f"实力表无危险品日期:\n{bad.to_string(index=False)}")

    end = pd.to_datetime(df["危险品"])
    start = end - pd.DateOffset(years=2)

    out = pd.DataFrame(
        {
            "参训人员姓名": df["姓名"],
            "工作单位": ORG_NAME,
            "统一社会信用代码": ORG_CREDIT,
            "工作单位类别": ORG_TYPE,
            "性别": df["性别"].astype(str).str.strip(),
            "证件类型": ID_TYPE,
            "证件号": df["员工号"],
            "手机号": df["手机号码"].map(lambda x: str(int(x)) if isinstance(x, float) else str(x).strip()),
            "国家二字码": df["国籍"].map(_nation_code),
            "培训类型": train_type,
            "考勤情况": DEFAULT_ATTEND,
            "课时": DEFAULT_HOURS,
            "随堂成绩": pd.to_numeric(df["态度评价"], errors="coerce"),
            "理论成绩": pd.to_numeric(df["考试"], errors="coerce"),
            "实操成绩": pd.to_numeric(df["任务模拟"], errors="coerce"),
            "总成绩": pd.to_numeric(df["评估得分"], errors="coerce"),
            "培训结果": df["评估结论"].astype(str).str.strip(),
            "有效期开始时间": start.dt.strftime("%Y-%m"),
            "有效期截止时间": end.dt.strftime("%Y-%m"),
            "备注": "",
        }
    )
    return out


def write_template(rows: pd.DataFrame, template: Path, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, output)
    wb = load_workbook(output)
    ws = wb["Sheet1"]
    cols = list(rows.columns)
    for i, row in enumerate(rows.itertuples(index=False), start=3):
        for j, val in enumerate(row, start=1):
            cell = ws.cell(i, j)
            if cols[j - 1] in {"证件号", "手机号", "统一社会信用代码"}:
                cell.value = str(val)
                cell.number_format = "@"
            elif cols[j - 1] in {"有效期开始时间", "有效期截止时间"}:
                cell.value = str(val)
                cell.number_format = "@"
            elif pd.isna(val) or val == "":
                cell.value = None
            else:
                cell.value = val
    wb.save(output)


def main() -> None:
    ap = argparse.ArgumentParser(description="危险品培训结果一次性导入填表")
    ap.add_argument("--zong", required=True, help="总成绩.xlsx")
    ap.add_argument("--caac", required=True, help="CAAC飞行员信息导出")
    ap.add_argument("--shili", required=True, help="飞行实力表（含人员信息.危险品）")
    ap.add_argument("--train-type", default="复训")
    ap.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    ap.add_argument("--out", required=True, help="输出 xlsx 路径")
    args = ap.parse_args()

    rows = build_rows(
        Path(args.zong),
        Path(args.caac),
        Path(args.shili),
        args.train_type,
    )
    write_template(rows, Path(args.template), Path(args.out))
    print(f"rows={len(rows)}")
    print(f"out={args.out}")
    print(rows.groupby(["有效期开始时间", "有效期截止时间"]).size().to_string())
    print("培训结果:", rows["培训结果"].value_counts().to_dict())


if __name__ == "__main__":
    main()
