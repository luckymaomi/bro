# -*- coding: utf-8 -*-
"""南货航人员结构报告：读实力表人员信息 → 八表统计与闭环 → 可选回填结构报告 docx。"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import zipfile
from collections import defaultdict
from copy import deepcopy
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}

REQUIRED_HEADERS = [
    "姓名",
    "技术信息",
    "RAMA",
    "REUO",
    "RWAS",
    "EAMA",
    "EEUO",
    "EWAS",
    "原单位",
    "检查员资格",
]

QUALIFICATION_CODES = [
    "RAMA",
    "REUO",
    "RWAS",
    "RSEA",
    "EAMA",
    "EEUO",
    "EWAS",
    "ESEA",
    "RANC",
    "RORD",
    "RJFK",
    "RLAX",
    "RNLU",
]

ORIGIN_LABELS = [
    "飞行/总队 777",
    "飞行/总队 737",
    "飞行/总队 320",
    "飞行/总队 909",
    "飞行/总队 919",
    "湖南",
    "湖北",
    "新疆",
    "河南",
    "西安",
    "重庆",
    "汕头",
    "珠海",
    "广西",
    "海南",
    "黑龙江",
    "上海",
    "顺丰航空",
]

SECTION_TITLES = [
    "教员、机长、副驾驶占比",
    "机长含以上各级别占比",
    "机长航线资格占比",
    "机长报务占比",
    "副驾驶级别占比",
    "副驾驶报务占比",
    "人员居住情况",
    "空勤人员原单位情况",
]

ALIASES = {
    "在训机长": "转机型机长",
    "在训副驾驶": "转机型副驾驶",
    "检查员（其中）": "检查员",
    "其中：检查员": "检查员",
    "777": "飞行/总队777",
    "737": "飞行/总队737",
    "320": "飞行/总队320",
    "909": "飞行/总队909",
}


@dataclass
class PersonnelRecord:
    employee_id: str
    name: str
    tech_info: str
    origin: str
    inspector_qualification: str
    qualifications: dict[str, bool]


@dataclass
class StatItem:
    label: str
    count: int
    denominator: int
    percent: str
    rule: str
    is_subset: bool = False


@dataclass
class StatSection:
    title: str
    denominator_label: str
    items: list[StatItem]
    total: int
    denominator: int
    closed: bool


@dataclass
class StructureResult:
    structure_crew_count: int
    captain_or_above_count: int
    first_officer_count: int
    sections: list[StatSection]
    warnings: list[str] = field(default_factory=list)
    unrecognized_tech: list[str] = field(default_factory=list)
    unrecognized_origin: list[str] = field(default_factory=list)

    @property
    def all_closed(self) -> bool:
        return all(section.closed for section in self.sections)


def _norm(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _norm_header(value: Any) -> str:
    return re.sub(r"\s+", "", _norm(value))


def _has_value(value: Any) -> bool:
    return _norm(value) != ""


def find_personnel_sheet(workbook) -> str:
    for name in workbook.sheetnames:
        if "人员信息" in name:
            return name
    raise SystemExit(f"未找到含「人员信息」的工作表，现有：{workbook.sheetnames}")


def parse_records(xlsx_path: Path, sheet_name: str | None = None) -> list[PersonnelRecord]:
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        sheet = sheet_name or find_personnel_sheet(wb)
        ws = wb[sheet]
        rows = [list(row) for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()

    header_idx = -1
    for i, row in enumerate(rows):
        headers = [_norm_header(c) for c in row]
        if sum(1 for h in REQUIRED_HEADERS if h in headers) >= 5:
            header_idx = i
            break
    if header_idx < 0:
        raise SystemExit("未识别到人员信息表头：" + "、".join(REQUIRED_HEADERS))

    header_map: dict[str, int] = {}
    for i, cell in enumerate(rows[header_idx]):
        key = _norm_header(cell)
        if key and key not in header_map:
            header_map[key] = i
    missing = [h for h in REQUIRED_HEADERS if h not in header_map]
    if missing:
        raise SystemExit("人员信息表缺少必要表头：" + "、".join(missing))

    records: list[PersonnelRecord] = []
    for row in rows[header_idx + 1 :]:
        def cell(header: str) -> Any:
            idx = header_map.get(header)
            return None if idx is None or idx >= len(row) else row[idx]

        name = _norm(cell("姓名"))
        tech = _norm(cell("技术信息"))
        emp = _norm(cell("员工号"))
        if not name and not tech and not emp:
            continue
        quals = {code: _has_value(cell(code)) for code in QUALIFICATION_CODES}
        records.append(
            PersonnelRecord(
                employee_id=emp,
                name=name,
                tech_info=tech,
                origin=_norm(cell("原单位")),
                inspector_qualification=_norm(cell("检查员资格")),
                qualifications=quals,
            )
        )
    return records


def tech_label(record: PersonnelRecord) -> str:
    parts = re.split(r"[:：]", record.tech_info, maxsplit=1)
    return _norm(parts[1] if len(parts) > 1 else record.tech_info)


def is_teacher(record: PersonnelRecord) -> bool:
    return "飞行教员" in tech_label(record)


def is_transfer_captain(record: PersonnelRecord) -> bool:
    return tech_label(record) == "划转机长"


def is_transfer_fo(record: PersonnelRecord) -> bool:
    return tech_label(record) == "划转副驾驶"


def is_captain(record: PersonnelRecord) -> bool:
    label = tech_label(record)
    return "机长" in label and "飞行教员" not in label and "划转" not in label


def is_regular_fo(record: PersonnelRecord) -> bool:
    label = tech_label(record)
    return "副驾驶" in label and "划转" not in label


def is_captain_or_above(record: PersonnelRecord) -> bool:
    return is_teacher(record) or is_captain(record) or is_transfer_captain(record)


def is_fo_group(record: PersonnelRecord) -> bool:
    return is_regular_fo(record) or is_transfer_fo(record)


def is_structure_crew(record: PersonnelRecord) -> bool:
    return (
        is_teacher(record)
        or is_captain(record)
        or is_transfer_captain(record)
        or is_regular_fo(record)
        or is_transfer_fo(record)
    )


def has_qual(record: PersonnelRecord, code: str) -> bool:
    return bool(record.qualifications.get(code))


def is_local(record: PersonnelRecord) -> bool:
    return record.origin.startswith("总队") or record.origin == "777返聘"


def is_inspector(record: PersonnelRecord) -> bool:
    return record.inspector_qualification in {"公司检查员", "委任代表"}


def is_line_captain(record: PersonnelRecord) -> bool:
    label = tech_label(record)
    captain_level = (
        "飞行教员" in label
        or "A类机长" in label
        or "B类机长" in label
        or "C类机长" in label
        or label == "E类机长"
        or label == "F类机长"
    )
    return (
        captain_level
        and "Z类机长" not in label
        and not has_qual(record, "RAMA")
        and not has_qual(record, "REUO")
        and not has_qual(record, "RWAS")
    )


def percent_round(count: int, denominator: int) -> str:
    if not denominator:
        return "0%"
    return f"{round(count / denominator * 100)}%"


def balance_percentages(items: list[StatItem], denominator: int) -> list[StatItem]:
    if not items or denominator <= 0:
        return [
            StatItem(i.label, i.count, i.denominator, "0%", i.rule, i.is_subset) for i in items
        ]
    exact = [i.count / denominator * 100 for i in items]
    allocated = [int(v) for v in exact]
    remainder = 100 - sum(allocated)
    order = sorted(
        range(len(exact)),
        key=lambda idx: (-(exact[idx] - allocated[idx]), idx),
    )
    for i in range(remainder):
        allocated[order[i % len(order)]] += 1
    out: list[StatItem] = []
    for item, pct in zip(items, allocated):
        out.append(
            StatItem(item.label, item.count, item.denominator, f"{pct}%", item.rule, item.is_subset)
        )
    return out


def make_item(
    label: str, count: int, denominator: int, rule: str, is_subset: bool = False
) -> StatItem:
    return StatItem(label, count, denominator, percent_round(count, denominator), rule, is_subset)


def count_if(records: list[PersonnelRecord], pred) -> int:
    return sum(1 for r in records if pred(r))


def format_people(records: list[PersonnelRecord]) -> str:
    parts = []
    for r in records:
        identity = " ".join(x for x in [r.employee_id, r.name] if x)
        parts.append(identity or tech_label(r))
    return "、".join(parts)


def make_closed_items(
    records: list[PersonnelRecord],
    denominator: int,
    definitions: list[tuple[str, Any, str]],
    other_prefix: str,
) -> list[StatItem]:
    matched: set[int] = set()
    items: list[StatItem] = []
    for label, pred, rule in definitions:
        matched_records = [r for r in records if pred(r)]
        for r in matched_records:
            matched.add(id(r))
        items.append(make_item(label, len(matched_records), denominator, rule))
    other = [r for r in records if id(r) not in matched]
    if other:
        items.append(
            make_item("其他", len(other), denominator, f"{other_prefix}：{format_people(other)}。")
        )
    return balance_percentages(items, denominator)


def make_combo_items(
    records: list[PersonnelRecord],
    denominator: int,
    prefix: str,
    labels: dict[str, str],
    rule_prefix: str,
) -> list[StatItem]:
    north = f"{prefix}AMA"
    europe = f"{prefix}EUO"
    west = f"{prefix}WAS"

    def combo(n: bool, e: bool, w: bool) -> int:
        return count_if(
            records,
            lambda r: has_qual(r, north) is n
            and has_qual(r, europe) is e
            and has_qual(r, west) is w,
        )

    return balance_percentages(
        [
            make_item("美+欧+西亚", combo(True, True, True), denominator, f"{rule_prefix}：同时具备北美、欧洲、西亚。"),
            make_item("美+欧", combo(True, True, False), denominator, f"{rule_prefix}：具备北美、欧洲，不具备西亚。"),
            make_item("美+西亚", combo(True, False, True), denominator, f"{rule_prefix}：具备北美、西亚，不具备欧洲。"),
            make_item("欧+西亚", combo(False, True, True), denominator, f"{rule_prefix}：具备欧洲、西亚，不具备北美。"),
            make_item(labels["northOnly"], combo(True, False, False), denominator, f"{rule_prefix}：只具备北美。"),
            make_item(labels["europeOnly"], combo(False, True, False), denominator, f"{rule_prefix}：只具备欧洲。"),
            make_item(labels["westOnly"], combo(False, False, True), denominator, f"{rule_prefix}：只具备西亚。"),
            make_item(labels["none"], combo(False, False, False), denominator, f"{rule_prefix}：北美、欧洲、西亚均不具备。"),
        ],
        denominator,
    )


def build_captain_route_items(records: list[PersonnelRecord], denominator: int) -> list[StatItem]:
    combo = make_combo_items(
        records,
        denominator,
        "R",
        {
            "northOnly": "仅北美带队",
            "europeOnly": "仅欧洲带队",
            "westOnly": "仅西亚带队",
            "none": "无美欧西亚单飞",
        },
        "RAMA/REUO/RWAS 单飞资格",
    )
    combo = [i for i in combo if i.label != "无美欧西亚单飞"]

    matched: set[int] = set()
    for r in records:
        has_any = has_qual(r, "RAMA") or has_qual(r, "REUO") or has_qual(r, "RWAS")
        if has_any or is_line_captain(r) or tech_label(r) == "Z类机长":
            matched.add(id(r))
    unmatched = [r for r in records if id(r) not in matched]
    route_count = count_if(records, is_line_captain) + len(unmatched)
    return balance_percentages(
        [
            *combo,
            make_item(
                "航线机长",
                route_count,
                denominator,
                "无RAMA/REUO/RWAS单飞资格且不是Z类机长的人员，统一归入航线机长；包含原本无法归入其他航线资格分类的人员。",
            ),
            make_item(
                "左座带飞",
                count_if(records, lambda r: tech_label(r) == "Z类机长"),
                denominator,
                "Z类机长。",
            ),
        ],
        denominator,
    )


def build_captain_level_items(records: list[PersonnelRecord], denominator: int) -> list[StatItem]:
    closed = make_closed_items(
        records,
        denominator,
        [
            ("C类教员", lambda r: tech_label(r) == "飞行教员C", "技术信息为飞行教员C。"),
            ("B类教员", lambda r: tech_label(r) == "飞行教员B", "技术信息为飞行教员B。"),
            ("F类机长", lambda r: tech_label(r) == "F类机长", "技术信息为F类机长。"),
            ("E类机长", lambda r: tech_label(r) == "E类机长", "技术信息为E类机长。"),
            ("D类机长", lambda r: tech_label(r) == "D类机长", "技术信息为D类机长。"),
            ("C类机长", lambda r: tech_label(r) == "C类机长", "技术信息为C类机长。"),
            ("B类机长", lambda r: tech_label(r) == "B类机长", "技术信息为B类机长。"),
            ("Z类机长", lambda r: tech_label(r) == "Z类机长", "技术信息为Z类机长。"),
            ("转机型机长", is_transfer_captain, "技术信息为划转机长，对外统一显示为转机型机长。"),
        ],
        "未落入教员/机长等级分类，需人工核对",
    )
    return [
        make_item(
            "检查员",
            count_if(records, is_inspector),
            denominator,
            "其中：检查员资格为公司检查员或委任代表，不参加技术等级构成求和。",
            True,
        ),
        *closed,
    ]


def build_fo_level_items(records: list[PersonnelRecord], denominator: int) -> list[StatItem]:
    return make_closed_items(
        records,
        denominator,
        [
            ("D类副驾驶", lambda r: tech_label(r) == "D类副驾驶", "技术信息为D类副驾驶。"),
            ("C类副驾驶", lambda r: tech_label(r) == "C类副驾驶", "技术信息为C类副驾驶。"),
            ("B类副驾驶", lambda r: tech_label(r) == "B类副驾驶", "技术信息为B类副驾驶。"),
            (
                "A类副驾驶",
                lambda r: tech_label(r) in {"A1类副驾驶", "A2类副驾驶"},
                "技术信息为A1类副驾驶或A2类副驾驶。",
            ),
            ("E类副驾驶", lambda r: tech_label(r) == "E类副驾驶", "技术信息为E类副驾驶。"),
            ("转机型副驾驶", is_transfer_fo, "技术信息为划转副驾驶，对外统一显示为转机型副驾驶。"),
        ],
        "未落入副驾驶等级分类，需人工核对",
    )


def map_origin(origin: str) -> str:
    normalized = _norm(origin)
    mapping = {
        "总队777": "飞行/总队 777",
        "777返聘": "飞行/总队 777",
        "总队737": "飞行/总队 737",
        "总队320": "飞行/总队 320",
        "总队909": "飞行/总队 909",
        "总队919": "飞行/总队 919",
        "湖南分公司": "湖南",
        "湖北分公司": "湖北",
        "新疆分公司": "新疆",
        "新疆分公司（借）": "新疆",
        "河南分公司": "河南",
        "西安分公司": "西安",
        "重庆航空": "重庆",
        "汕头分公司": "汕头",
        "珠海分公司": "珠海",
        "广西分公司": "广西",
        "海南分公司": "海南",
        "黑龙江分公司": "黑龙江",
        "上海分公司": "上海",
        "上海分公司（借）": "上海",
        "顺丰航空": "顺丰航空",
    }
    return mapping.get(normalized, normalized or "未识别")


def make_section(
    title: str, denominator_label: str, items: list[StatItem], closure_denominator: int
) -> StatSection:
    total = sum(i.count for i in items if not i.is_subset)
    return StatSection(
        title=title,
        denominator_label=denominator_label,
        items=items,
        total=total,
        denominator=closure_denominator,
        closed=total == closure_denominator,
    )


def calculate(records: list[PersonnelRecord]) -> StructureResult:
    structure_crew = [r for r in records if is_structure_crew(r)]
    captain_base = [r for r in records if is_teacher(r) or is_captain(r)]
    captain_or_above = [r for r in records if is_captain_or_above(r)]
    fo_base = [r for r in records if is_regular_fo(r)]
    fo_with_transfer = [r for r in records if is_fo_group(r)]

    sc = len(structure_crew)
    cap_den = len(captain_base)
    cao = len(captain_or_above)
    fo_den = len(fo_base)
    fo_n = len(fo_with_transfer)

    sections = [
        make_section(
            "教员、机长、副驾驶占比",
            f"{sc}人",
            balance_percentages(
                [
                    make_item("教员", count_if(structure_crew, is_teacher), sc, "技术信息包含飞行教员。"),
                    make_item(
                        "机长",
                        count_if(structure_crew, lambda r: is_captain(r) or is_transfer_captain(r)),
                        sc,
                        "非教员机长，含转机型机长。",
                    ),
                    make_item(
                        "副驾驶",
                        count_if(structure_crew, lambda r: is_regular_fo(r) or is_transfer_fo(r)),
                        sc,
                        "副驾驶，含转机型副驾驶。",
                    ),
                ],
                sc,
            ),
            sc,
        ),
        make_section(
            "机长含以上各级别占比",
            f"{cao}人",
            build_captain_level_items(captain_or_above, cao),
            cao,
        ),
        make_section(
            "机长航线资格占比",
            f"{cap_den}人，不含转机型",
            build_captain_route_items(captain_base, cap_den),
            cap_den,
        ),
        make_section(
            "机长报务占比",
            f"{cap_den}人，不含转机型",
            make_combo_items(
                captain_base,
                cap_den,
                "E",
                {
                    "northOnly": "单美洲报务",
                    "europeOnly": "单欧洲报务",
                    "westOnly": "单西亚报务",
                    "none": "无报务",
                },
                "EAMA/EEUO/EWAS 英语通信资格",
            ),
            cap_den,
        ),
        make_section(
            "副驾驶级别占比",
            f"{fo_n}人",
            build_fo_level_items(fo_with_transfer, fo_n),
            fo_n,
        ),
        make_section(
            "副驾驶报务占比",
            f"{fo_den}人，不含转机型",
            make_combo_items(
                fo_base,
                fo_den,
                "E",
                {
                    "northOnly": "单美洲报务",
                    "europeOnly": "单欧洲报务",
                    "westOnly": "单西亚报务",
                    "none": "无报务",
                },
                "EAMA/EEUO/EWAS 英语通信资格",
            ),
            fo_den,
        ),
        make_section(
            "人员居住情况",
            f"{cao}人 / {fo_n}人",
            [
                *balance_percentages(
                    [
                        make_item(
                            "机长本地居住",
                            count_if(captain_or_above, is_local),
                            cao,
                            "原单位以总队开头或等于777返聘。",
                        ),
                        make_item(
                            "机长异地居住",
                            count_if(captain_or_above, lambda r: not is_local(r)),
                            cao,
                            "除本地外均为异地。",
                        ),
                    ],
                    cao,
                ),
                *balance_percentages(
                    [
                        make_item(
                            "副驾驶本地居住",
                            count_if(fo_with_transfer, is_local),
                            fo_n,
                            "原单位以总队开头或等于777返聘。",
                        ),
                        make_item(
                            "副驾驶异地居住",
                            count_if(fo_with_transfer, lambda r: not is_local(r)),
                            fo_n,
                            "除本地外均为异地。",
                        ),
                    ],
                    fo_n,
                ),
            ],
            sc,
        ),
    ]

    origin_counts: dict[str, int] = defaultdict(int)
    origin_people: dict[str, list[PersonnelRecord]] = defaultdict(list)
    for r in structure_crew:
        label = map_origin(r.origin)
        origin_counts[label] += 1
        origin_people[label].append(r)

    other_entries = [(k, v) for k, v in origin_counts.items() if k not in ORIGIN_LABELS]
    other_count = sum(v for _, v in other_entries)
    other_rule = ""
    if other_entries:
        bits = [
            f"{label}（{format_people(origin_people[label])}）" for label, _ in other_entries
        ]
        other_rule = "未映射原单位，需人工核对：" + "；".join(bits) + "。"

    origin_items = [
        make_item(
            label,
            origin_counts.get(label, 0),
            sc,
            "按原单位映射汇总；总队777与777返聘合并到飞行/总队777，总队919归入飞行/总队919。",
        )
        for label in ORIGIN_LABELS
    ]
    if other_count:
        origin_items.append(make_item("其他", other_count, sc, other_rule))
    sections.append(
        make_section(
            "空勤人员原单位情况",
            f"{sc}人",
            balance_percentages(origin_items, sc),
            sc,
        )
    )

    recognized = [
        r
        for r in records
        if is_teacher(r)
        or is_captain(r)
        or is_regular_fo(r)
        or is_transfer_captain(r)
        or is_transfer_fo(r)
    ]
    recognized_ids = {id(r) for r in recognized}
    unrecognized_tech = sorted(
        {
            r.tech_info
            for r in records
            if r.tech_info and id(r) not in recognized_ids
        }
    )
    unrecognized_origin = sorted(
        {
            r.origin
            for r in structure_crew
            if r.origin and map_origin(r.origin) not in ORIGIN_LABELS
        }
    )

    warnings: list[str] = []
    if unrecognized_tech:
        warnings.append(f"有 {len(unrecognized_tech)} 类技术信息未识别。")
    if unrecognized_origin:
        warnings.append(f"有 {len(unrecognized_origin)} 类原单位未映射。")
    for section in sections:
        for item in section.items:
            if item.label == "其他" and item.count > 0:
                warnings.append(f"{section.title}：其他 {item.count} 人，{item.rule}")
        if not section.closed:
            warnings.append(
                f"{section.title}：构成合计 {section.total} 人，与母数 {section.denominator} 人不一致。"
            )

    return StructureResult(
        structure_crew_count=sc,
        captain_or_above_count=cao,
        first_officer_count=fo_n,
        sections=sections,
        warnings=warnings,
        unrecognized_tech=unrecognized_tech,
        unrecognized_origin=unrecognized_origin,
    )


def export_excel(result: StructureResult, out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "统计结果"
    ws.append(["分组/表", "构成项", "人数", "占比", "母数", "其中项", "口径"])
    for section in result.sections:
        for item in section.items:
            ws.append(
                [
                    section.title,
                    item.label,
                    item.count,
                    item.percent,
                    item.denominator,
                    "是" if item.is_subset else "",
                    item.rule,
                ]
            )

    ws2 = wb.create_sheet("闭环核对")
    ws2.append(["表", "构成合计", "母数", "闭环", "母数说明"])
    for section in result.sections:
        ws2.append(
            [
                section.title,
                section.total,
                section.denominator,
                "是" if section.closed else "否",
                section.denominator_label,
            ]
        )
    ws2.append([])
    ws2.append(["结构统计人员", result.structure_crew_count])
    ws2.append(["机长含以上", result.captain_or_above_count])
    ws2.append(["副驾驶", result.first_officer_count])
    ws2.append(["全部闭环", "是" if result.all_closed else "否"])

    ws3 = wb.create_sheet("规则说明")
    ws3.append(["说明"])
    ws3.append(["飞行管理人员、地面人员不进结构统计母数。"])
    ws3.append(["是否运行、行政职务不参与筛选。"])
    ws3.append(["检查员为其中项，不进等级构成求和。"])
    ws3.append(["划转对外显示为转机型；航线／报务母数不含转机型（副驾报务不含转机型副驾驶）。"])
    ws3.append(["本月报口径独立于挖数成长口径；E类副驾驶在本报告副驾级别表中单列。"])

    ws4 = wb.create_sheet("未识别数据")
    ws4.append(["类型", "原值"])
    for value in result.unrecognized_tech:
        ws4.append(["技术信息", value])
    for value in result.unrecognized_origin:
        ws4.append(["原单位", value])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def _qn(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _word_text(node: ET.Element) -> str:
    parts: list[str] = []
    for el in node.iter():
        if el.tag == _qn("t") and el.text:
            parts.append(el.text)
        elif el.tag in {_qn("br"), _qn("tab"), _qn("cr")}:
            parts.append(" ")
    return "".join(parts)


def _norm_word(node: ET.Element) -> str:
    return re.sub(r"\s+", "", _word_text(node))


def _first_paragraph(node: ET.Element) -> ET.Element | None:
    if node.tag == _qn("p"):
        return node
    return node.find(_qn("p"))


def _clone_style_bits(source: ET.Element | None) -> tuple[ET.Element | None, ET.Element | None]:
    """从单元格／段落取出 pPr、首个 rPr，用于回填时保住对齐字体。"""
    if source is None:
        return None, None
    paragraph = _first_paragraph(source)
    if paragraph is None:
        return None, None
    p_pr = paragraph.find(_qn("pPr"))
    r_pr = None
    first_run = paragraph.find(_qn("r"))
    if first_run is not None:
        r_pr = first_run.find(_qn("rPr"))
    return (
        deepcopy(p_pr) if p_pr is not None else None,
        deepcopy(r_pr) if r_pr is not None else None,
    )


def _set_word_text(target: ET.Element, value: str, style_from: ET.Element | None = None) -> None:
    """只改文字；保留 tcPr，并尽量沿用本格或参照格的段落／字体格式。"""
    p_pr, r_pr = _clone_style_bits(target)
    if style_from is not None:
        donor_p, donor_r = _clone_style_bits(style_from)
        if p_pr is None and donor_p is not None:
            p_pr = donor_p
        if r_pr is None and donor_r is not None:
            r_pr = donor_r

    keep: list[ET.Element] = []
    if target.tag == _qn("tc"):
        tc_pr = target.find(_qn("tcPr"))
        if tc_pr is not None:
            keep.append(tc_pr)
    for child in list(target):
        target.remove(child)
    for el in keep:
        target.append(el)

    p = ET.SubElement(target, _qn("p"))
    if p_pr is not None:
        p.append(p_pr)
    r = ET.SubElement(p, _qn("r"))
    if r_pr is not None:
        r.append(r_pr)
    t = ET.SubElement(r, _qn("t"))
    t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = value


def _style_donor_for_row(row: dict[str, Any], prefer_col: int | None, month_cols: list[int]) -> ET.Element | None:
    """同行走样式：优先上月，再往前找有内容的月份格。"""
    cells = row["cells"]
    order: list[int] = []
    if prefer_col is not None:
        order.append(prefer_col)
    for col in reversed(month_cols):
        if col not in order:
            order.append(col)
    for col in order:
        if col < 0 or col >= len(cells):
            continue
        cell = cells[col]
        if cell is None:
            continue
        if _norm_word(cell):
            return cell
        # 空格也可能带居中 pPr，仍可作样式源
        p_pr, r_pr = _clone_style_bits(cell)
        if p_pr is not None or r_pr is not None:
            return cell
    return None


def _word_rows(table: ET.Element) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for tr in table.findall(_qn("tr")):
        cells: list[ET.Element] = []
        for tc in tr.findall(_qn("tc")):
            props = tc.find(_qn("tcPr"))
            span_el = props.find(_qn("gridSpan")) if props is not None else None
            span = int(span_el.get(f"{{{W_NS}}}val", "1")) if span_el is not None else 1
            merge = props.find(_qn("vMerge")) if props is not None else None
            continued = False
            if merge is not None:
                val = merge.get(f"{{{W_NS}}}val")
                continued = val != "restart"
            previous = result[-1]["cells"][len(cells)] if result and len(result[-1]["cells"]) > len(cells) else None
            if continued and previous is None:
                raise SystemExit("Word 表格纵向合并没有起始单元格。")
            for _ in range(span):
                cells.append(previous if continued else tc)
        result.append({"element": tr, "cells": cells})
    return result


def _table_columns(table: ET.Element) -> dict[str, Any]:
    rows = _word_rows(table)
    if not rows:
        raise SystemExit("Word 空表。")
    header = rows[0]["cells"]
    months: dict[int, int] = {}
    change = percent = -1
    for index, cell in enumerate(header):
        text = _norm_word(cell)
        m = re.fullmatch(r"(\d{1,2})月", text)
        if m:
            month = int(m.group(1))
            if month < 1 or month > 12 or month in months:
                raise SystemExit("Word 月份表头无效或重复。")
            months[month] = index
        if text == "本月变化":
            change = index
        if text == "本月占比":
            percent = index
    if not months or change < 0 or percent < 0:
        raise SystemExit("Word 表格需包含月份、本月变化和本月占比列。")
    return {"months": months, "change": change, "percent": percent}


def _label_column(index: int) -> int:
    return 1 if index in {1, 6, 7} else 0


def _label_for(row: dict[str, Any], index: int) -> str:
    cells = row["cells"]
    label = _norm_word(cells[_label_column(index)])
    if index == 6:
        group = _norm_word(cells[0])
        prefix = "副驾驶" if "副驾驶" in group else "机长" if "机长" in group else ""
        return f"{prefix}{label}"
    if index == 7 and "飞行总队" in _norm_word(cells[0]):
        return f"飞行/总队{label}"
    return ALIASES.get(label, label)


def _group_for(label: str, index: int, result: StructureResult) -> str:
    if index == 1:
        teachers = next(i.count for i in result.sections[0].items if i.label == "教员")
        if label in {"C类教员", "B类教员"}:
            return f"教员（{teachers}）"
        if label == "检查员":
            return ""
        return f"机长（{result.captain_or_above_count - teachers}）"
    if index == 6:
        return (
            f"机长（{result.captain_or_above_count}）"
            if label.startswith("机长")
            else f"副驾驶（{result.first_officer_count}）"
        )
    if index == 7:
        return "飞行总队" if label.startswith("飞行/总队 ") else label
    return ""


def _change_text(count: int, previous: str) -> str:
    if not re.fullmatch(r"\d+", previous or ""):
        return "/"
    diff = count - int(previous)
    if diff > 0:
        return f"新增{diff}人"
    if diff < 0:
        return f"减少{-diff}人"
    return "/"


def _tc_grid_span(tc: ET.Element) -> int:
    props = tc.find(_qn("tcPr"))
    if props is None:
        return 1
    span_el = props.find(_qn("gridSpan"))
    if span_el is None:
        return 1
    try:
        return int(span_el.get(f"{{{W_NS}}}val", "1"))
    except ValueError:
        return 1


def _clone_row_for_append(
    table: ET.Element,
    label_column: int,
    label: str,
    group: str,
    *,
    origin_mode: str | None = None,
) -> None:
    """补分类行。

    原单位表固化两种模板：
    - fleet：飞行总队 + 机型 双列（gridSpan=1+1）
    - single：上海／顺丰／黑龙江等名称跨两列（gridSpan=2），只写一格，避免「黑龙江|黑龙江」
    其他表仍按 label_column／group 旧逻辑。
    """
    rows = _word_rows(table)
    template = None
    if origin_mode == "single":
        for row in reversed(rows[1:]):
            physical = row["element"].findall(_qn("tc"))
            if physical and _tc_grid_span(physical[0]) >= 2:
                template = row
                break
    elif origin_mode == "fleet":
        for row in reversed(rows[1:]):
            physical = row["element"].findall(_qn("tc"))
            if (
                len(physical) >= 2
                and _tc_grid_span(physical[0]) == 1
                and "飞行总队" in _norm_word(physical[0])
            ):
                template = row
                break
    if template is None:
        for row in reversed(rows[1:]):
            cells = row["cells"]
            if cells[label_column] is not None and (
                label_column == 0 or cells[0] is not cells[label_column]
            ):
                template = row
                break
    if template is None:
        raise SystemExit("Word 表格没有可复制的数据行。")

    new_tr = deepcopy(template["element"])
    style_seed = None
    physical_src = template["element"].findall(_qn("tc"))
    if physical_src:
        style_seed = physical_src[0]
    for tc in new_tr.findall(_qn("tc")):
        props = tc.find(_qn("tcPr"))
        if props is not None:
            for merge in list(props.findall(_qn("vMerge"))):
                props.remove(merge)
        _set_word_text(tc, "", style_from=style_seed)
    table.append(new_tr)

    if origin_mode == "single":
        first = new_tr.findall(_qn("tc"))[0]
        _set_word_text(first, label, style_from=style_seed)
        return

    added = _word_rows(table)[-1]
    if origin_mode == "fleet":
        _set_word_text(added["cells"][0], "飞行总队", style_from=added["cells"][0])
        _set_word_text(added["cells"][1], label, style_from=added["cells"][1])
        return

    if label_column > 0:
        _set_word_text(added["cells"][0], group, style_from=added["cells"][0])
    _set_word_text(added["cells"][label_column], label, style_from=added["cells"][label_column])


def inspect_docx_months(docx_path: Path) -> dict[str, Any]:
    with zipfile.ZipFile(docx_path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find(_qn("body"))
    if body is None:
        raise SystemExit("Word 未找到正文。")
    tables = body.findall(_qn("tbl"))
    if len(tables) < 9:
        raise SystemExit("Word 人员结构模板至少需要 9 张统计表。")
    target_tables = tables[1:9]
    common: set[int] | None = None
    for table in target_tables:
        months = set(_table_columns(table)["months"])
        common = months if common is None else common & months
    if not common:
        raise SystemExit("八张 Word 统计表没有共同月份。")
    first_rows = _word_rows(target_tables[0])[1:]
    labels = {_norm_word(r["cells"][0]) for r in first_rows}
    if not {"教员", "机长", "副驾驶"}.issubset(labels):
        raise SystemExit("Word 人员占比表缺少教员、机长或副驾驶分类。")
    empty = []
    cols = _table_columns(target_tables[0])
    for month in sorted(common):
        col = cols["months"][month]
        if all(_norm_word(r["cells"][col]) == "" for r in first_rows):
            empty.append(month)
    return {
        "months": sorted(common),
        "suggested_month": empty[0] if len(empty) == 1 else None,
    }


def _read_int_cell(text: str) -> int | None:
    raw = _norm(text)
    return int(raw) if re.fullmatch(r"\d+", raw or "") else None


def read_previous_headcount(docx_path: Path, month: int) -> dict[str, Any]:
    """从模板读上月管理人员、文首地面人数，供默认沿用。"""
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(_qn("body"))
    if body is None:
        raise SystemExit("Word 未找到正文。")
    tables = body.findall(_qn("tbl"))
    if not tables:
        raise SystemExit("Word 无表。")
    table0 = tables[0]
    cols = _table_columns(table0)
    prev_month = month - 1
    prev_col = cols["months"].get(prev_month)
    managers_prev = None
    for row in _word_rows(table0)[1:]:
        label = _norm_word(row["cells"][0])
        if label == "管理人员" and prev_col is not None:
            managers_prev = _read_int_cell(_norm_word(row["cells"][prev_col]))
            break
    ground_prev = None
    aircrew_prev = None
    for p in body.findall(_qn("p")):
        text = _word_text(p)
        if "总人数及空地人员占比" not in text:
            continue
        m_air = re.search(r"空勤人员（已注册人员）(\d+)人", text)
        m_ground = re.search(r"地面人员(\d+)人", text)
        if m_air:
            aircrew_prev = int(m_air.group(1))
        if m_ground:
            ground_prev = int(m_ground.group(1))
        break
    return {
        "managers_prev": managers_prev,
        "ground_prev": ground_prev if ground_prev is not None else 30,
        "aircrew_prev": aircrew_prev,
    }


def _rewrite_paragraph(el: ET.Element, text: str) -> None:
    p_pr, r_pr = _clone_style_bits(el)
    for child in list(el):
        el.remove(child)
    if p_pr is not None:
        el.append(p_pr)
    r = ET.SubElement(el, _qn("r"))
    if r_pr is not None:
        r.append(r_pr)
    t = ET.SubElement(r, _qn("t"))
    t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text


def _fill_headcount_and_managers(
    body: ET.Element,
    tables: list[ET.Element],
    month: int,
    aircrew: int,
    ground: int,
    managers: int,
    warnings: list[str],
) -> None:
    if managers < 0 or ground < 0:
        raise SystemExit("管理人员／地面人员人数不能为负。")
    if managers > aircrew:
        raise SystemExit(f"管理人员 {managers} 大于已注册空勤 {aircrew}，无法填第1张表。")
    non_managers = aircrew - managers
    total = aircrew + ground
    pct_air = round(aircrew / total * 100) if total else 0
    pct_ground = 100 - pct_air if total else 0
    intro = (
        f"总人数及空地人员占比：总人数{total}人，含空勤人员（已注册人员）{aircrew}人、"
        f"地面人员{ground}人，空勤人员占{pct_air}%、地面人员占{pct_ground}%。"
    )

    changed_bits = []
    for p in body.findall(_qn("p")):
        text = _word_text(p)
        if "总人数及空地人员占比" in text:
            _rewrite_paragraph(p, intro)
            warnings.append(f"已更新文首：{intro}")
        elif _norm_word(p).startswith("飞行管理人员占比"):
            _rewrite_paragraph(p, f"飞行管理人员占比（{aircrew}人）")
            warnings.append(f"已更新标题：飞行管理人员占比（{aircrew}人）")

    table0 = tables[0]
    cols = _table_columns(table0)
    target = cols["months"].get(month)
    if target is None:
        raise SystemExit(f"管理人员表未找到 {month}月 列。")
    previous = cols["months"].get(month - 1)
    counts = {"管理人员": managers, "非管理人员": non_managers}
    pct_mgr = round(managers / aircrew * 100) if aircrew else 0
    pct_non = 100 - pct_mgr if aircrew else 0
    percents = {"管理人员": f"{pct_mgr}%", "非管理人员": f"{pct_non}%"}
    month_cols = list(cols["months"].values())
    for row in _word_rows(table0)[1:]:
        label = _norm_word(row["cells"][0])
        if label not in counts:
            continue
        count = counts[label]
        prev_text = _norm_word(row["cells"][previous]) if previous is not None else ""
        prev_n = _read_int_cell(prev_text)
        if prev_n is not None and prev_n != count:
            changed_bits.append(label)
        donor = _style_donor_for_row(row, previous, month_cols)
        _set_word_text(row["cells"][target], str(count), style_from=donor)
        _set_word_text(
            row["cells"][cols["change"]],
            _change_text(count, prev_text),
            style_from=row["cells"][cols["change"]] if _norm_word(row["cells"][cols["change"]]) else donor,
        )
        _set_word_text(
            row["cells"][cols["percent"]],
            percents[label],
            style_from=row["cells"][cols["percent"]] if _norm_word(row["cells"][cols["percent"]]) else donor,
        )
    warnings.append(
        f"第1张表已写 {month}月：管理人员{managers}、非管理人员{non_managers}（空勤{aircrew}）。"
    )
    if changed_bits:
        warnings.append(
            "管理人员相对上月有变化：" + "、".join(changed_bits) + "；文首「发生变化」勾选请主人过目。"
        )


def fill_docx(
    template_path: Path,
    result: StructureResult,
    month: int,
    out_path: Path,
    *,
    aircrew: int,
    ground: int,
    managers: int,
) -> list[str]:
    if not result.all_closed:
        raise SystemExit("统计未闭环，不能填充 Word。")
    if month < 1 or month > 12:
        raise SystemExit("写入月份须为 1–12。")

    tmp_dir = out_path.parent / f"_docx_fill_{out_path.stem}"
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    tmp_dir.mkdir(parents=True)
    with zipfile.ZipFile(template_path) as zf:
        zf.extractall(tmp_dir)

    doc_xml = tmp_dir / "word" / "document.xml"
    # Preserve namespaces as much as possible
    ET.register_namespace("w", W_NS)
    tree = ET.parse(doc_xml)
    root = tree.getroot()
    body = root.find(_qn("body"))
    assert body is not None
    tables = body.findall(_qn("tbl"))
    if len(tables) < 9:
        raise SystemExit("Word 人员结构模板至少需要 9 张统计表。")

    warnings = list(result.warnings)
    filled_rows = 0
    _fill_headcount_and_managers(body, tables, month, aircrew, ground, managers, warnings)

    # Find title paragraphs immediately before tables 2..9
    children = list(body)
    table_positions = [i for i, el in enumerate(children) if el.tag == _qn("tbl")]
    if len(table_positions) < 9:
        raise SystemExit("Word 表数量不足。")

    for section_index, title in enumerate(SECTION_TITLES):
        table = tables[section_index + 1]
        section = next(s for s in result.sections if s.title == title)

        # title paragraph
        tbl_pos = table_positions[section_index + 1]
        title_el = None
        for j in range(tbl_pos - 1, -1, -1):
            el = children[j]
            if el.tag == _qn("tbl"):
                break
            if el.tag == _qn("p") and _norm_word(el).startswith(title):
                title_el = el
                break
        if title_el is None:
            raise SystemExit(f"Word 第{section_index + 2}张表未找到对应标题：{title}。")

        # remove A类教员
        if section_index == 1:
            for row in list(_word_rows(table)[1:]):
                if _label_for(row, section_index) == "A类教员":
                    table.remove(row["element"])
                    warnings.append(f"{title}：删除 A类教员 行，不推算历史 E/F 人数。")

        existing = {_label_for(row, section_index) for row in _word_rows(table)[1:]}
        for item in section.items:
            key = re.sub(r"\s+", "", item.label)
            if key in {re.sub(r"\s+", "", x) for x in existing}:
                continue
            if item.count == 0 and item.label not in {"E类机长", "F类机长"}:
                continue
            label = item.label
            origin_mode = None
            if section_index == 6:
                label = re.sub(r"^(机长|副驾驶)", "", item.label)
            elif section_index == 7:
                if item.label.startswith("飞行/总队"):
                    label = re.sub(r"^飞行/总队\s*", "", item.label)
                    origin_mode = "fleet"
                else:
                    label = item.label
                    origin_mode = "single"
            _clone_row_for_append(
                table,
                _label_column(section_index),
                label,
                _group_for(item.label, section_index, result),
                origin_mode=origin_mode,
            )
            existing.add(item.label)
            warnings.append(f"{title}：模板补充分类行 {item.label}。")

        config = _table_columns(table)
        target = config["months"].get(month)
        if target is None:
            raise SystemExit(f"{title} 未找到 {month}月 列。")
        previous = config["months"].get(month - 1)
        month_cols = list(config["months"].values())
        items = {re.sub(r"\s+", "", item.label): item for item in section.items}
        mapped: set[str] = set()
        groups: dict[ET.Element, str] = {}
        written: set[ET.Element] = set()

        for row in _word_rows(table)[1:]:
            key = _label_for(row, section_index)
            item = items.get(re.sub(r"\s+", "", key))
            prev_text = _norm_word(row["cells"][previous]) if previous is not None else ""
            donor = _style_donor_for_row(row, previous, month_cols)
            change_cell = row["cells"][config["change"]]
            percent_cell = row["cells"][config["percent"]]
            target_cell = row["cells"][target]
            output_cells = [target_cell, change_cell, percent_cell]
            for cell in output_cells:
                written.add(cell)
            if not item:
                _set_word_text(target_cell, "", style_from=donor)
                _set_word_text(change_cell, "", style_from=change_cell if _norm_word(change_cell) else donor)
                _set_word_text(percent_cell, "", style_from=percent_cell if _norm_word(percent_cell) else donor)
                warnings.append(f"{title}：未识别模板分类 {key or '空白'}，本月列留空。")
                continue
            if key in mapped:
                raise SystemExit(f"{title} 模板分类重复：{item.label}。")
            mapped.add(key)
            if previous is not None and not re.fullmatch(r"\d+", prev_text or ""):
                warnings.append(f"{title} / {item.label}：上月无有效人数，本月变化为“/”。")
            _set_word_text(target_cell, str(item.count), style_from=donor)
            _set_word_text(
                change_cell,
                _change_text(item.count, prev_text),
                style_from=change_cell if _norm_word(change_cell) else donor,
            )
            _set_word_text(
                percent_cell,
                item.percent,
                style_from=percent_cell if _norm_word(percent_cell) else donor,
            )
            if section_index in {1, 6} and not item.is_subset:
                group = _group_for(item.label, section_index, result)
                groups[row["cells"][0]] = group
            if section_index == 1 and _norm_word(row["cells"][1]) in ALIASES:
                _set_word_text(row["cells"][1], item.label, style_from=row["cells"][1])
            if section_index == 4 and _norm_word(row["cells"][0]) in ALIASES:
                _set_word_text(row["cells"][0], item.label, style_from=row["cells"][0])
            filled_rows += 1

        for cell, group in groups.items():
            _set_word_text(cell, group, style_from=cell)

        excluded = "不含转机型" if section_index in {2, 3, 5} else ""
        # rewrite title paragraph text roughly
        for child in list(title_el):
            title_el.remove(child)
        r = ET.SubElement(title_el, _qn("r"))
        t = ET.SubElement(r, _qn("t"))
        t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = f"{title}（{section.denominator}人{excluded}）"

    warnings.append("人员引进及其他部门仍保持模板原文，请人工核对。")
    tree.write(doc_xml, encoding="UTF-8", xml_declaration=True)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in tmp_dir.rglob("*"):
            if path.is_file():
                zf.write(path, path.relative_to(tmp_dir).as_posix())
    shutil.rmtree(tmp_dir)
    warnings.append(f"已写入 {filled_rows} 行数据到 {month} 月列。")
    return warnings


def result_to_dict(result: StructureResult) -> dict[str, Any]:
    return {
        "structureCrewCount": result.structure_crew_count,
        "captainOrAboveCount": result.captain_or_above_count,
        "firstOfficerCount": result.first_officer_count,
        "allClosed": result.all_closed,
        "warnings": result.warnings,
        "unrecognized": {
            "techInfo": result.unrecognized_tech,
            "origin": result.unrecognized_origin,
        },
        "sections": [
            {
                "title": s.title,
                "denominatorLabel": s.denominator_label,
                "total": s.total,
                "denominator": s.denominator,
                "closed": s.closed,
                "items": [asdict(i) for i in s.items],
            }
            for s in result.sections
        ],
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="南货航人员结构报告统计与 Word 回填")
    ap.add_argument("--xlsx", required=True, help="飞行实力表路径")
    ap.add_argument("--sheet", default="", help="人员信息工作表名；省略则自动找含「人员信息」的表")
    ap.add_argument("--docx", default="", help="结构报告 Word 模板；提供则回填")
    ap.add_argument("--month", type=int, default=0, help="写入月份 1-12；0=仅预检/看建议月")
    ap.add_argument("--ground", type=int, default=-1, help="地面人员人数；-1=沿用模板文首(常见30)")
    ap.add_argument("--managers", type=int, default=-1, help="管理人员人数；-1=沿用上月管理人员列")
    ap.add_argument("--out-xlsx", default="", help="统计 Excel 输出路径")
    ap.add_argument("--out-docx", default="", help="回填 Word 输出路径")
    ap.add_argument("--out-json", default="", help="统计 JSON 输出路径")
    ap.add_argument("--inspect-only", action="store_true", help="只检查 Word 月份，不算表")
    args = ap.parse_args()

    xlsx = Path(args.xlsx)
    if not xlsx.exists():
        raise SystemExit(f"找不到实力表：{xlsx}")

    prev_info: dict[str, Any] | None = None
    if args.docx:
        docx = Path(args.docx)
        if not docx.exists():
            raise SystemExit(f"找不到 Word：{docx}")
        info = inspect_docx_months(docx)
        print("Word月份:", ",".join(str(m) for m in info["months"]))
        print("建议写入月:", info["suggested_month"] if info["suggested_month"] else "无唯一空列，请主人指定")
        if args.inspect_only:
            return
        if args.month:
            prev_info = read_previous_headcount(docx, args.month)
            print(
                f"模板默认：上月管理人员={prev_info['managers_prev']} "
                f"文首地面={prev_info['ground_prev']} 文首空勤旧值={prev_info['aircrew_prev']}"
            )

    records = parse_records(xlsx, args.sheet or None)
    aircrew = len(records)
    result = calculate(records)
    print(
        f"已注册空勤(表内人数)={aircrew} 结构统计人员={result.structure_crew_count} "
        f"机长含以上={result.captain_or_above_count} 副驾驶={result.first_officer_count} "
        f"闭环={'是' if result.all_closed else '否'}"
    )
    for section in result.sections:
        flag = "闭环" if section.closed else "未闭环"
        print(f"- {section.title}: 合计{section.total}/母数{section.denominator} [{flag}]")
    for w in result.warnings:
        print("警告:", w)

    if args.out_json:
        path = Path(args.out_json)
        path.parent.mkdir(parents=True, exist_ok=True)
        payload = result_to_dict(result)
        payload["registeredAircrew"] = aircrew
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print("JSON:", path)
    if args.out_xlsx:
        export_excel(result, Path(args.out_xlsx))
        print("Excel:", args.out_xlsx)

    if args.docx and args.month:
        assert prev_info is not None
        ground = args.ground if args.ground >= 0 else int(prev_info["ground_prev"])
        if args.managers >= 0:
            managers = args.managers
        else:
            if prev_info["managers_prev"] is None:
                raise SystemExit("模板上月无管理人员人数，请 --managers 明示。")
            managers = int(prev_info["managers_prev"])
        print(f"采用：空勤{aircrew} 地面{ground} 管理人员{managers} 总人数{aircrew + ground}")
        out_docx = Path(args.out_docx) if args.out_docx else Path(f"南货航人员结构报告_{args.month:02d}月_回填.docx")
        warns = fill_docx(
            Path(args.docx),
            result,
            args.month,
            out_docx,
            aircrew=aircrew,
            ground=ground,
            managers=managers,
        )
        print("Word:", out_docx)
        for w in warns:
            print("Word提示:", w)
    elif args.docx and not args.month:
        print("已算表；未指定 --month，未回填 Word。开场仍须问：管理人员变了吗？地面人员变了吗？")


if __name__ == "__main__":
    main()
