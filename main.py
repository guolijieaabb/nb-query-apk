# -*- coding: utf-8 -*-
"""
NB--综合查询工具（安卓手机简化版）
功能：设备信息查询 + 维保查询。已删除申请单/Word 生成/维修类型选择。
仅依赖 kivy + openpyxl，便于打包成安卓 APK。
"""
import os
import sys
import openpyxl

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.core.window import Window

EXCEL_FILE = "2026090400441.xlsx"
WBAO_FILE = "weibao.xlsx"
SEARCH_COLS = ["管理IP", "序列号", "固资编号", "位置", "业务地址"]
DETAIL_FIELDS = ["设备分类", "采购用途", "管理IP", "业务地址", "部署网络", "状态",
                 "设备类型", "设备型号", "序列号", "数据中心", "位置", "建设工程", "固资编号"]


def resource_path(filename: str) -> str:
    """定位数据文件：打包/运行后，优先脚本(私有)目录下的同名文件。"""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


def load_sheet(path):
    """用 openpyxl 读整张表，返回 list[dict]，键为表头。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = []
    for r in rows[1:]:
        d = {}
        for i, h in enumerate(headers):
            d[h] = r[i] if i < len(r) else ""
        data.append(d)
    return data


def find_col(columns, *keywords):
    for c in columns:
        cn = str(c).replace(" ", "").upper()
        for kw in keywords:
            if kw.upper().replace(" ", "") in cn:
                return c
    return None


class QueryApp(App):
    def build(self):
        self.title = "NB综合查询"
        self.rows = []        # 主表全部
        self.results = []     # 查询结果
        self.idx = 0

        root = BoxLayout(orientation="vertical", padding=10, spacing=6)

        # 顶部：搜索框 + 按钮
        top = BoxLayout(size_hint_y=None, height=48, spacing=6)
        self.kw = TextInput(multiline=False, hint_text="输入 IP/序列号/固资编号/位置/业务地址",
                            size_hint_x=0.6, font_size=16)
        self.kw.bind(on_text_validate=lambda x: self.do_search())
        top.add_widget(self.kw)
        top.add_widget(Button(text="查询", on_release=lambda x: self.do_search(), size_hint_x=0.15))
        top.add_widget(Button(text="清空", on_release=lambda x: self.clear(), size_hint_x=0.12))
        top.add_widget(Button(text="维保", on_release=lambda x: self.on_weibao(), size_hint_x=0.15))
        root.add_widget(top)

        # 状态栏
        self.status = Label(text="正在加载...", size_hint_y=None, height=24, color=(0.2, 0.2, 0.2, 1),
                            font_size=13)
        root.add_widget(self.status)

        # 设备详情
        self.detail = Label(text="", valign="top", halign="left", size_hint_y=1,
                            text_size=(Window.width - 24, None), font_size=15)
        sv = ScrollView()
        sv.add_widget(self.detail)
        root.add_widget(sv)

        # 底部：翻页
        bottom = BoxLayout(size_hint_y=None, height=48, spacing=6)
        bottom.add_widget(Button(text="上一条", on_release=lambda x: self.prev()))
        self.page = Label(text="0 / 0", size_hint_x=0.3)
        bottom.add_widget(self.page)
        bottom.add_widget(Button(text="下一条", on_release=lambda x: self.next()))
        root.add_widget(bottom)

        # 加载数据
        self.load_excel()
        return root

    def load_excel(self):
        p = resource_path(EXCEL_FILE)
        if not os.path.exists(p):
            self.status.text = f"找不到 {EXCEL_FILE}"
            return
        try:
            self.rows = load_sheet(p)
            self.status.text = f"已加载 {len(self.rows)} 条设备"
        except Exception as e:
            self.status.text = f"读取失败：{e}"

    def do_search(self):
        kw = self.kw.text.strip().lower()
        self.results = []
        self.idx = 0
        if not kw:
            self.status.text = "请输入关键词"
            self.detail.text = ""
            self.page.text = "0 / 0"
            return
        for row in self.rows:
            for col in SEARCH_COLS:
                if col in row and kw in str(row[col]).lower():
                    self.results.append(row)
                    break
        self.status.text = f"查询到 {len(self.results)} 条"
        if self.results:
            self.show_current()
        else:
            self.detail.text = "无匹配结果"
            self.page.text = "0 / 0"

    def show_current(self):
        row = self.results[self.idx]
        lines = [f"{k}：{row.get(k, '')}" for k in DETAIL_FIELDS]
        self.detail.text = "\n".join(lines)
        self.page.text = f"{self.idx + 1} / {len(self.results)}"

    def clear(self):
        self.kw.text = ""
        self.results = []
        self.idx = 0
        self.detail.text = ""
        self.page.text = "0 / 0"
        self.status.text = "已清空"

    def prev(self):
        if self.idx > 0:
            self.idx -= 1
            self.show_current()

    def next(self):
        if self.idx < len(self.results) - 1:
            self.idx += 1
            self.show_current()

    def on_weibao(self):
        if not self.results:
            self.popup("提示", "请先查询出一条设备，再点【维保】")
            return
        cur = self.results[self.idx]
        targets = set()
        for v in (cur.get("序列号"), cur.get("固资编号"), cur.get("位置")):
            if v is not None and str(v).strip().lower() not in ("", "nan"):
                targets.add(str(v).strip().lower())
        if self.kw.text.strip():
            targets.add(self.kw.text.strip().lower())
        wp = resource_path(WBAO_FILE)
        if not os.path.exists(wp):
            self.popup("维保查询", f"找不到 {WBAO_FILE}")
            return
        try:
            wdata = load_sheet(wp)
        except Exception as e:
            self.popup("维保查询", f"读取失败：{e}")
            return
        cols = list(wdata[0].keys()) if wdata else []
        sn_col = find_col(cols, "S/N", "SN", "序列号")
        loc_col = find_col(cols, "安装地点", "安装位置", "位置")
        asset_col = find_col(cols, "固资编号")
        month_col = find_col(cols, "维保期限", "剩余月数", "月数")
        match_cols = [c for c in (sn_col, loc_col, asset_col) if c]
        if not match_cols:
            self.popup("维保查询", f"未识别到 S/N/安装地点/固资编号 列")
            return
        hit = None
        for row in wdata:
            for c in match_cols:
                val = str(row.get(c, "")).strip().lower()
                if val and val in targets:
                    hit = row
                    break
            if hit:
                break
        if not hit:
            self.popup("维保查询", "❌ 该设备不在华创维保范围内")
            return
        if month_col:
            msg = f"✅ 在华创维保范围内\n剩余维保期限：{hit.get(month_col)} 个月"
        else:
            msg = "✅ 在华创维保范围内（未识别到月数列）"
        self.popup("维保查询", msg)

    def popup(self, title, text):
        lbl = Label(text=text, valign="top", halign="left", size_hint_y=None)
        lbl.bind(texture_size=lambda w, v: setattr(lbl, "height", v[1]))
        sv = ScrollView(size_hint=(1, 1))
        sv.add_widget(lbl)
        box = BoxLayout(orientation="vertical", padding=8)
        box.add_widget(sv)
        box.add_widget(Button(text="关闭", size_hint_y=None, height=48,
                              on_release=lambda x: p.dismiss()))
        p = Popup(title=title, content=box, size_hint=(0.9, 0.6))
        p.open()


if __name__ == "__main__":
    QueryApp().run()
