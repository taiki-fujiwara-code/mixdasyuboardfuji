"""
グループ クチコミ管理 一括返信・自動返信 画面仕様書
GMO TECH 4:3 ブルーテンプレに準拠したPowerPoint資料を生成
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from copy import deepcopy
from lxml import etree

# ===== カラー定義（GMO TECHブランドカラーに準拠） =====
BRAND_BLUE      = RGBColor(0x1F, 0x4E, 0x99)   # メイン
DEEP_NAVY       = RGBColor(0x0C, 0x1E, 0x3C)   # ダークアクセント
LIGHT_BLUE      = RGBColor(0xCB, 0xDC, 0xF0)   # 補助
ACCENT_ORANGE   = RGBColor(0xEE, 0x82, 0x1F)   # 強調
TEXT_BLACK      = RGBColor(0x21, 0x21, 0x21)
TEXT_GRAY       = RGBColor(0x66, 0x66, 0x66)
LINE_GRAY       = RGBColor(0xCC, 0xCC, 0xCC)
BG_LIGHT        = RGBColor(0xF5, 0xF6, 0xF8)
BG_HEAD         = RGBColor(0xE9, 0xEE, 0xF6)
WHITE           = RGBColor(0xFF, 0xFF, 0xFF)
WARN_BG         = RGBColor(0xFF, 0xF3, 0xCD)
SUCCESS_GREEN   = RGBColor(0x2E, 0x8B, 0x57)

# 4:3 寸法
SLIDE_W = Inches(10.0)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

BLANK = prs.slide_layouts[6]


# ===== ヘルパー =====
def add_textbox(slide, x, y, w, h, text, *,
                font_size=18, bold=False, color=TEXT_BLACK,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                font_name="Noto Sans JP", line_spacing=1.3):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = ln
        r.font.name = font_name
        r.font.size = Pt(font_size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def add_rect(slide, x, y, w, h, fill, line=None, line_w=0.75, shadow=False):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    if not shadow:
        # 影を除去
        sp = s.shadow
        sp.inherit = False
    return s


def add_rounded_rect(slide, x, y, w, h, fill, line=None, line_w=0.75):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.adjustments[0] = 0.10
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s


def add_parallelogram(slide, x, y, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, x, y, w, h)
    s.adjustments[0] = 0.35
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def add_right_triangle(slide, x, y, w, h, fill, flip_h=False, flip_v=False):
    s = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    # flipH / flipV を spPr/xfrm 直下に正しく設定
    sp = s._element
    spPr_list = sp.findall(".//" + qn("p:spPr"))
    if spPr_list:
        spPr = spPr_list[0]
        xfrm = spPr.find(qn("a:xfrm"))
        if xfrm is not None:
            if flip_h:
                xfrm.set("flipH", "1")
            if flip_v:
                xfrm.set("flipV", "1")
    return s


def add_line(slide, x1, y1, x2, y2, color=LINE_GRAY, weight=0.75, dash=False):
    ln = slide.shapes.add_connector(1, x1, y1, x2, y2)
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    if dash:
        lnEl = ln.line._get_or_add_ln()
        prstDash = etree.SubElement(lnEl, qn("a:prstDash"))
        prstDash.set("val", "dash")
    return ln


# ===== マスター装飾 =====
def add_master_decor(slide, page_num):
    """全ページ共通: 左上の青パラレログラム + 右上GMO TECH ロゴ + ページ番号"""
    # 左上青三角（タイトルバー）
    add_parallelogram(slide, Inches(-0.05), Inches(0.18), Inches(1.05), Inches(0.45), BRAND_BLUE)

    # 右上 GMO TECH ロゴ（テキストで再現）
    logo = slide.shapes.add_textbox(Inches(8.45), Inches(0.32), Inches(1.5), Inches(0.34))
    tf = logo.text_frame
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r1 = p.add_run(); r1.text = "GMO"
    r1.font.name = "Noto Sans JP"; r1.font.size = Pt(18); r1.font.bold = True
    r1.font.color.rgb = BRAND_BLUE
    r2 = p.add_run(); r2.text = "TECH"
    r2.font.name = "Noto Sans JP"; r2.font.size = Pt(18); r2.font.bold = True
    r2.font.color.rgb = TEXT_GRAY

    # ページ番号（右下）
    pn = slide.shapes.add_textbox(Inches(9.4), Inches(7.1), Inches(0.5), Inches(0.3))
    tf = pn.text_frame
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run(); r.text = str(page_num)
    r.font.name = "Noto Sans JP"; r.font.size = Pt(11); r.font.color.rgb = TEXT_GRAY


def add_slide_title(slide, title_text):
    """ページタイトル（左上の青パラレログラムの右側に配置）"""
    add_textbox(slide, Inches(0.55), Inches(0.20), Inches(8.0), Inches(0.5),
                title_text, font_size=24, bold=True, color=TEXT_BLACK)


# ===== Slide 1: タイトル =====
def slide_title():
    s = prs.slides.add_slide(BLANK)
    # 背景: 左斜めのダークネイビー（画像なしで再現）
    # 大きな白背景の上に左下〜中央にかけて暗色の三角
    add_right_triangle(s, Inches(-1.0), Inches(0), Inches(6.5), Inches(7.5),
                       DEEP_NAVY, flip_h=True)
    # 右上の青三角
    add_right_triangle(s, Inches(7.0), Inches(-1.0), Inches(4.0), Inches(4.0),
                       BRAND_BLUE)
    # 左下の小さな青三角
    add_right_triangle(s, Inches(-0.5), Inches(6.4), Inches(1.7), Inches(1.5),
                       BRAND_BLUE, flip_v=True)

    # タイトルテキスト（白）
    add_textbox(s, Inches(1.0), Inches(2.6), Inches(7.5), Inches(0.6),
                "PRODUCT SPEC — GROUP REVIEW MANAGEMENT",
                font_size=12, bold=True, color=WHITE)
    add_textbox(s, Inches(1.0), Inches(3.05), Inches(8.0), Inches(0.9),
                "グループ クチコミ管理機能",
                font_size=32, bold=True, color=WHITE)
    add_textbox(s, Inches(1.0), Inches(3.85), Inches(8.0), Inches(0.7),
                "一括返信・自動返信 画面設計仕様書",
                font_size=24, bold=True, color=WHITE)

    # 補助情報
    add_textbox(s, Inches(1.0), Inches(5.5), Inches(7.0), Inches(0.4),
                "Version 1.0  /  2026.06.15",
                font_size=12, color=LIGHT_BLUE)

    # GMO TECH ロゴ（右下）
    logo = s.shapes.add_textbox(Inches(7.2), Inches(6.7), Inches(2.5), Inches(0.5))
    tf = logo.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    r1 = p.add_run(); r1.text = "GMO"
    r1.font.name = "Noto Sans JP"; r1.font.size = Pt(20); r1.font.bold = True
    r1.font.color.rgb = WHITE
    r2 = p.add_run(); r2.text = "TECH"
    r2.font.name = "Noto Sans JP"; r2.font.size = Pt(20); r2.font.bold = True
    r2.font.color.rgb = LIGHT_BLUE

    # フッターCopyright
    add_textbox(s, Inches(5.0), Inches(7.1), Inches(5.0), Inches(0.3),
                "Copyright © GMO TECH All Rights reserved.   Confidential",
                font_size=10, color=WHITE, align=PP_ALIGN.RIGHT)


# ===== Slide 2: 目次 =====
def slide_toc():
    s = prs.slides.add_slide(BLANK)
    add_master_decor(s, 2)

    # 左側 大きなダークブロック（目次用）
    add_rect(s, Inches(0), Inches(0.95), Inches(3.4), Inches(6.2), DEEP_NAVY)
    add_textbox(s, Inches(0.5), Inches(4.0), Inches(2.5), Inches(0.6),
                "目次", font_size=32, bold=True, color=WHITE)
    add_textbox(s, Inches(0.5), Inches(4.7), Inches(2.5), Inches(0.4),
                "CONTENTS", font_size=12, color=LIGHT_BLUE)

    # 右側 項目
    items = [
        ("1.", "背景と本機能の目的"),
        ("2.", "現状の課題（単店舗UIとの比較）"),
        ("3.", "画面全体構成 / 遷移"),
        ("4.", "【画面A】一括返信 UI設計"),
        ("5.", "【画面B】自動返信ルール設定 UI設計"),
        ("6.", "返信テンプレート管理"),
        ("7.", "設定項目・データモデル"),
        ("8.", "権限・実装フェーズ / 要確認事項"),
    ]
    base_y = 1.3
    for i, (no, label) in enumerate(items):
        y = Inches(base_y + i * 0.62)
        add_textbox(s, Inches(3.85), y, Inches(0.6), Inches(0.45),
                    no, font_size=22, bold=True, color=BRAND_BLUE)
        add_textbox(s, Inches(4.5), y, Inches(5.3), Inches(0.45),
                    label, font_size=18, color=TEXT_BLACK)


# ===== Slide 3: 背景と目的 =====
def slide_purpose():
    s = prs.slides.add_slide(BLANK)
    add_master_decor(s, 3)
    add_slide_title(s, "1. 背景と本機能の目的")

    # 小見出し
    add_textbox(s, Inches(0.5), Inches(1.05), Inches(9.0), Inches(0.45),
                "本機能の目的", font_size=20, bold=True, color=BRAND_BLUE)

    # 説明ボックス
    add_rounded_rect(s, Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.45), BG_LIGHT)
    add_textbox(s, Inches(0.75), Inches(1.70), Inches(8.6), Inches(1.3),
                "グループ管理機能において、クチコミの「一括返信」「自動返信ルール」設定を、\n"
                "現在のCSV/Excelインポート運用ではなくツール上で直接GUI操作で完結できるようにする。\n"
                "単店舗で既に提供されている「クチコミ返信パターン・自動返信」画面のUIを継承し、\n"
                "グループ横断の「適用店舗」概念のみを差分追加することで、学習コストを最小化する。",
                font_size=15, color=TEXT_BLACK, line_spacing=1.35)

    # 小見出し
    add_textbox(s, Inches(0.5), Inches(3.20), Inches(9.0), Inches(0.45),
                "期待される効果", font_size=20, bold=True, color=BRAND_BLUE)

    # 3カラムの効果
    cols = [
        ("運用効率の向上", "CSVのDL/編集/UL作業が不要になり、\n設定完了までの工数を1/5程度に短縮"),
        ("ヒューマンエラー削減", "テンプレ・条件をGUIでバリデーション\nしながら設定でき、入力ミスが減少"),
        ("グループ横断の運用", "複数店舗に対し同一ルールを一括で\n適用でき、ブランド統制が容易"),
    ]
    col_w = 3.0
    for i, (h, b) in enumerate(cols):
        x = Inches(0.5 + i * 3.05)
        add_rect(s, x, Inches(3.75), Inches(col_w), Inches(0.55), BRAND_BLUE)
        add_textbox(s, x, Inches(3.83), Inches(col_w), Inches(0.4),
                    h, font_size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_rect(s, x, Inches(4.30), Inches(col_w), Inches(1.55), WHITE, line=LINE_GRAY)
        add_textbox(s, x + Inches(0.15), Inches(4.45), Inches(col_w - 0.3), Inches(1.4),
                    b, font_size=14, color=TEXT_BLACK)

    # 対象ユーザー
    add_textbox(s, Inches(0.5), Inches(6.05), Inches(9.0), Inches(0.4),
                "対象ユーザー", font_size=20, bold=True, color=BRAND_BLUE)
    add_textbox(s, Inches(0.5), Inches(6.55), Inches(9.0), Inches(0.5),
                "・グループ管理者（master / full）：ルール作成・編集・削除・一括返信実行\n"
                "・編集ユーザー（edit）：一括返信のみ実行可、ルールは閲覧のみ\n"
                "・閲覧ユーザー（view）：参照のみ",
                font_size=14, color=TEXT_BLACK, line_spacing=1.35)


# ===== Slide 4: 現状の課題 =====
def slide_gap():
    s = prs.slides.add_slide(BLANK)
    add_master_decor(s, 4)
    add_slide_title(s, "2. 現状の課題（単店舗UIとの比較）")

    # 注釈バー
    add_rect(s, Inches(0.5), Inches(1.0), Inches(9.0), Inches(0.5), BG_HEAD)
    add_textbox(s, Inches(0.7), Inches(1.07), Inches(8.7), Inches(0.4),
                "※単店舗には「クチコミ返信パターン・自動返信」画面が既に存在。グループにも同等GUIを提供する。",
                font_size=13, color=TEXT_BLACK)

    # 比較表ヘッダー
    headers = ["観点", "単店舗（現状）", "グループ（現状）", "グループ（本案）"]
    col_xs = [0.5, 1.9, 4.1, 6.7]
    col_ws = [1.4, 2.2, 2.6, 2.8]
    head_y = Inches(1.7)
    for i, h in enumerate(headers):
        add_rect(s, Inches(col_xs[i]), head_y, Inches(col_ws[i]), Inches(0.45), BRAND_BLUE)
        add_textbox(s, Inches(col_xs[i]), Inches(1.78), Inches(col_ws[i]), Inches(0.35),
                    h, font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    rows = [
        ("返信パターン登録", "GUIで20件まで登録可", "CSVインポートのみ", "GUI登録 + 適用店舗指定"),
        ("自動返信ON/OFF", "パターン単位でチェックボックス", "店舗毎にCSV運用", "ルール単位でトグルスイッチ"),
        ("適用店舗の指定", "当店舗のみ（概念なし）", "CSV列で指定（不可視）", "全店舗 / 個別選択（明示UI）"),
        ("条件設定", "感情・★数・コメント有無", "同上（CSVで定義）", "感情・★数・キーワード等 拡張"),
        ("一括返信実行", "—", "—（インポート前提）", "絞込→選択→送信のStep導線"),
        ("操作の取消", "登録前なら可", "再インポート要", "都度キャンセル/編集/削除可"),
    ]
    row_h = 0.55
    for i, row in enumerate(rows):
        y = Inches(2.15 + i * row_h)
        bg = WHITE if i % 2 == 0 else BG_LIGHT
        for j, cell in enumerate(row):
            add_rect(s, Inches(col_xs[j]), y, Inches(col_ws[j]), Inches(row_h),
                     bg, line=LINE_GRAY)
            color = ACCENT_ORANGE if j == 3 else TEXT_BLACK
            bold  = True if j == 3 else False
            size  = 11 if j != 0 else 12
            add_textbox(s, Inches(col_xs[j] + 0.1), y + Inches(0.10),
                        Inches(col_ws[j] - 0.2), Inches(row_h - 0.15),
                        cell, font_size=size, bold=bold, color=color)


# ===== Slide 5: 画面全体構成 =====
def slide_structure():
    s = prs.slides.add_slide(BLANK)
    add_master_decor(s, 5)
    add_slide_title(s, "3. 画面全体構成 / タブ遷移")

    # 小見出し
    add_textbox(s, Inches(0.5), Inches(1.05), Inches(9.0), Inches(0.4),
                "既存「クチコミ管理」配下に2タブを新規追加", font_size=20, bold=True, color=BRAND_BLUE)

    # パンくず
    add_textbox(s, Inches(0.5), Inches(1.6), Inches(9.0), Inches(0.4),
                "グループ管理  ＞  〇〇グループ  ＞  クチコミ管理",
                font_size=13, color=TEXT_GRAY)

    # タブ表現
    tabs = [
        ("クチコミ一覧", "既存", False),
        ("一括返信", "新規", True),
        ("自動返信ルール", "新規", True),
        ("返信テンプレート", "新規", True),
    ]
    x = 0.5
    for label, badge, is_new in tabs:
        w = 2.25
        bg = WHITE if not is_new else BRAND_BLUE
        fg = TEXT_BLACK if not is_new else WHITE
        add_rect(s, Inches(x), Inches(2.10), Inches(w), Inches(0.55),
                 bg, line=LINE_GRAY if not is_new else BRAND_BLUE)
        add_textbox(s, Inches(x), Inches(2.20), Inches(w), Inches(0.35),
                    label, font_size=14, bold=True, color=fg, align=PP_ALIGN.CENTER)
        # NEW バッジ
        if is_new:
            add_rounded_rect(s, Inches(x + w - 0.7), Inches(2.13),
                             Inches(0.55), Inches(0.22), ACCENT_ORANGE)
            add_textbox(s, Inches(x + w - 0.7), Inches(2.13),
                        Inches(0.55), Inches(0.22),
                        "NEW", font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        x += w + 0.05

    # 区切り
    add_rect(s, Inches(0.5), Inches(2.65), Inches(9.0), Inches(0.04), BRAND_BLUE)

    # 各タブの説明カード
    add_textbox(s, Inches(0.5), Inches(2.95), Inches(9.0), Inches(0.4),
                "各タブの役割", font_size=18, bold=True, color=BRAND_BLUE)

    cards = [
        ("一括返信", "未返信クチコミを絞り込み→対象を選択→\n返信文を入力 or テンプレ選択 → 一括送信。\n大量件数は Sidekiq で非同期処理。",
         "Step UI（3段階）"),
        ("自動返信ルール", "★数・キーワード・適用店舗ごとに条件設定。\n新着クチコミに対し定期実行で自動返信。\nルールはトグルでON/OFF、優先度並べ替え。",
         "単店舗UIをグループ向けに拡張"),
        ("返信テンプレート", "本部承認済みの返信文をテンプレ化し、\n一括返信・自動返信の両方から呼び出し可能。\n本部での文言統制が可能。",
         "他タブから参照される共通資産"),
    ]
    for i, (head, body, badge) in enumerate(cards):
        y = 3.45 + i * 1.20
        add_rect(s, Inches(0.5), Inches(y), Inches(2.0), Inches(1.10), BRAND_BLUE)
        add_textbox(s, Inches(0.5), Inches(y + 0.35), Inches(2.0), Inches(0.4),
                    head, font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_rect(s, Inches(2.5), Inches(y), Inches(5.4), Inches(1.10),
                 WHITE, line=LINE_GRAY)
        add_textbox(s, Inches(2.65), Inches(y + 0.10), Inches(5.2), Inches(1.0),
                    body, font_size=13, color=TEXT_BLACK, line_spacing=1.3)
        add_rect(s, Inches(7.9), Inches(y), Inches(1.6), Inches(1.10),
                 BG_LIGHT, line=LINE_GRAY)
        add_textbox(s, Inches(7.95), Inches(y + 0.35), Inches(1.5), Inches(0.45),
                    badge, font_size=11, bold=True, color=BRAND_BLUE,
                    align=PP_ALIGN.CENTER)


# ===== Slide 6: 一括返信 UI =====
def slide_bulk_reply_ui():
    s = prs.slides.add_slide(BLANK)
    add_master_decor(s, 6)
    add_slide_title(s, "4. 【画面A】一括返信 UI設計")

    # 注釈
    add_textbox(s, Inches(0.5), Inches(0.95), Inches(9.0), Inches(0.4),
                "Step1: 絞り込み → Step2: クチコミ選択 → Step3: 返信文設定 の3ステップ構成",
                font_size=14, color=TEXT_GRAY)

    # ===== ワイヤーフレーム枠 =====
    wf_x, wf_y = 0.5, 1.45
    wf_w, wf_h = 9.0, 5.45
    add_rect(s, Inches(wf_x), Inches(wf_y), Inches(wf_w), Inches(wf_h),
             WHITE, line=BRAND_BLUE, line_w=1.2)
    # ヘッダーバー
    add_rect(s, Inches(wf_x), Inches(wf_y), Inches(wf_w), Inches(0.40), BRAND_BLUE)
    add_textbox(s, Inches(wf_x + 0.15), Inches(wf_y + 0.05),
                Inches(wf_w - 0.3), Inches(0.32),
                "クチコミ管理 > 一括返信",
                font_size=12, bold=True, color=WHITE)

    # Step1 ボックス
    s1y = wf_y + 0.55
    add_rect(s, Inches(wf_x + 0.2), Inches(s1y),
             Inches(wf_w - 0.4), Inches(1.10), BG_LIGHT, line=LINE_GRAY)
    add_textbox(s, Inches(wf_x + 0.35), Inches(s1y + 0.05),
                Inches(2.5), Inches(0.3),
                "■ Step1  クチコミを絞り込む", font_size=12, bold=True, color=BRAND_BLUE)
    # フィルタ入力（プレース表現）
    f_items = ["店舗 [すべて▼]", "評価 ☑★1 ☑★2 ☐★3 ☑★4 ☑★5",
               "ステータス [未返信▼]", "期間 [yyyy/mm/dd〜yyyy/mm/dd]"]
    for i, t in enumerate(f_items):
        x = wf_x + 0.35 + (i % 2) * 4.3
        y = s1y + 0.40 + (i // 2) * 0.32
        add_textbox(s, Inches(x), Inches(y), Inches(4.0), Inches(0.30),
                    t, font_size=11, color=TEXT_BLACK)

    # Step2 テーブル
    s2y = s1y + 1.20
    add_rect(s, Inches(wf_x + 0.2), Inches(s2y),
             Inches(wf_w - 0.4), Inches(2.40), WHITE, line=LINE_GRAY)
    add_textbox(s, Inches(wf_x + 0.35), Inches(s2y + 0.05),
                Inches(5.0), Inches(0.3),
                "■ Step2  返信するクチコミを選択する",
                font_size=12, bold=True, color=BRAND_BLUE)
    add_textbox(s, Inches(wf_x + wf_w - 2.2), Inches(s2y + 0.05),
                Inches(2.0), Inches(0.3),
                "選択中: 12件 / 全32件",
                font_size=11, color=TEXT_GRAY, align=PP_ALIGN.RIGHT)
    # テーブル
    cols = ["☐", "店舗名", "投稿日", "評価", "クチコミ本文（抜粋）", "状態"]
    col_w = [0.4, 1.4, 1.0, 0.8, 3.5, 0.9]
    tx = wf_x + 0.35
    ty = s2y + 0.42
    # ヘッダー
    add_rect(s, Inches(tx), Inches(ty), Inches(sum(col_w)), Inches(0.32), BG_HEAD)
    cx = tx
    for i, c in enumerate(cols):
        add_textbox(s, Inches(cx + 0.04), Inches(ty + 0.04),
                    Inches(col_w[i] - 0.08), Inches(0.24),
                    c, font_size=10, bold=True, color=TEXT_BLACK)
        cx += col_w[i]
    # 行
    sample = [
        ("☑", "渋谷店",  "06/10", "★★★★☆", "スタッフの対応がとても丁寧でした…", "未返信"),
        ("☑", "新宿店",  "06/09", "★★★☆☆", "駐車場が分かりにくかったです",       "未返信"),
        ("☐", "池袋店",  "06/09", "★★★★★", "とても良かったです、また行きます",   "未返信"),
        ("☑", "立川店",  "06/08", "★★☆☆☆", "待ち時間が長く感じた",              "未返信"),
        ("☐", "町田店",  "06/07", "★★★★★", "素晴らしい接客でした",              "返信済"),
    ]
    for r_i, row in enumerate(sample):
        ry = ty + 0.32 + r_i * 0.30
        bg = WHITE if r_i % 2 == 0 else BG_LIGHT
        add_rect(s, Inches(tx), Inches(ry), Inches(sum(col_w)), Inches(0.30),
                 bg, line=LINE_GRAY)
        cx = tx
        for i, c in enumerate(row):
            col = SUCCESS_GREEN if (i == 5 and c == "返信済") else \
                  ACCENT_ORANGE if (i == 5) else TEXT_BLACK
            add_textbox(s, Inches(cx + 0.04), Inches(ry + 0.03),
                        Inches(col_w[i] - 0.08), Inches(0.24),
                        c, font_size=9, color=col)
            cx += col_w[i]

    # Step3 返信文設定
    s3y = s2y + 2.50
    add_rect(s, Inches(wf_x + 0.2), Inches(s3y),
             Inches(wf_w - 0.4), Inches(0.95), BG_LIGHT, line=LINE_GRAY)
    add_textbox(s, Inches(wf_x + 0.35), Inches(s3y + 0.05),
                Inches(5.0), Inches(0.3),
                "■ Step3  返信文を設定する",
                font_size=12, bold=True, color=BRAND_BLUE)
    add_textbox(s, Inches(wf_x + 0.35), Inches(s3y + 0.35),
                Inches(5.0), Inches(0.3),
                "返信モード  ● 共通文を入力   ○ テンプレートから選択",
                font_size=10, color=TEXT_BLACK)
    add_rect(s, Inches(wf_x + 0.35), Inches(s3y + 0.62),
             Inches(6.0), Inches(0.28), WHITE, line=LINE_GRAY)
    add_textbox(s, Inches(wf_x + 0.40), Inches(s3y + 0.64),
                Inches(5.9), Inches(0.24),
                "［返信文テキストエリア／最大300文字］",
                font_size=10, color=TEXT_GRAY)
    add_rounded_rect(s, Inches(wf_x + 6.55), Inches(s3y + 0.60),
                     Inches(1.85), Inches(0.32), ACCENT_ORANGE)
    add_textbox(s, Inches(wf_x + 6.55), Inches(s3y + 0.62),
                Inches(1.85), Inches(0.28),
                "🤖 AIで返信文生成",
                font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # 実行ボタン
    by = s3y + 1.05
    add_textbox(s, Inches(wf_x + 0.35), Inches(by + 0.05),
                Inches(5.0), Inches(0.3),
                "⚠ 選択中の3件のクチコミに返信します。この操作は取り消せません。",
                font_size=10, color=ACCENT_ORANGE)
    add_rounded_rect(s, Inches(wf_x + 7.0), Inches(by),
                     Inches(1.4), Inches(0.32), BRAND_BLUE)
    add_textbox(s, Inches(wf_x + 7.0), Inches(by + 0.03),
                Inches(1.4), Inches(0.28),
                "一括返信を実行",
                font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ===== Slide 7: 一括返信 機能詳細 =====
def slide_bulk_reply_detail():
    s = prs.slides.add_slide(BLANK)
    add_master_decor(s, 7)
    add_slide_title(s, "4. 【画面A】一括返信 機能詳細")

    # 左：絞り込み条件
    add_textbox(s, Inches(0.5), Inches(1.0), Inches(4.5), Inches(0.4),
                "絞り込み条件", font_size=18, bold=True, color=BRAND_BLUE)
    # 表
    rows_l = [
        ("店舗", "ドロップダウン（複数選択可）", "すべて"),
        ("評価", "チェックボックス ★1-★5", "すべて"),
        ("返信ステータス", "すべて / 未返信 / 返信済", "未返信"),
        ("期間", "投稿日 from-to", "過去30日"),
        ("キーワード", "クチコミ本文 部分一致", "（空）"),
    ]
    hx = [0.5, 1.6, 3.5]
    hw = [1.1, 1.9, 1.05]
    ty = 1.45
    headers = ["項目", "入力種別", "初期値"]
    for i, h in enumerate(headers):
        add_rect(s, Inches(hx[i]), Inches(ty), Inches(hw[i]), Inches(0.35), BRAND_BLUE)
        add_textbox(s, Inches(hx[i]), Inches(ty + 0.03),
                    Inches(hw[i]), Inches(0.3), h,
                    font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for r_i, row in enumerate(rows_l):
        y = ty + 0.35 + r_i * 0.42
        bg = WHITE if r_i % 2 == 0 else BG_LIGHT
        for j, c in enumerate(row):
            add_rect(s, Inches(hx[j]), Inches(y), Inches(hw[j]), Inches(0.42),
                     bg, line=LINE_GRAY)
            add_textbox(s, Inches(hx[j] + 0.05), Inches(y + 0.06),
                        Inches(hw[j] - 0.1), Inches(0.32),
                        c, font_size=10, color=TEXT_BLACK)

    # 右：返信モードと実行
    add_textbox(s, Inches(5.2), Inches(1.0), Inches(4.5), Inches(0.4),
                "返信モード", font_size=18, bold=True, color=BRAND_BLUE)
    modes = [
        ("共通文を入力", "全選択クチコミに同一テキストを返信。\n300文字制限、文字数カウンタ表示。"),
        ("テンプレートから選択", "登録済みテンプレを呼び出し。\n選択後の編集も可能。"),
        ("🤖 AIで返信文を生成", "トーン・文字数・追加指示を指定。\n複数件は個別生成オプション有。"),
    ]
    for i, (h, b) in enumerate(modes):
        y = 1.45 + i * 1.18
        add_rect(s, Inches(5.2), Inches(y), Inches(0.18), Inches(1.05), BRAND_BLUE)
        add_rect(s, Inches(5.38), Inches(y), Inches(4.32), Inches(1.05),
                 WHITE, line=LINE_GRAY)
        add_textbox(s, Inches(5.55), Inches(y + 0.10), Inches(4.0), Inches(0.32),
                    h, font_size=14, bold=True, color=TEXT_BLACK)
        add_textbox(s, Inches(5.55), Inches(y + 0.45), Inches(4.0), Inches(0.6),
                    b, font_size=11, color=TEXT_GRAY)

    # 下部：実行フロー
    add_textbox(s, Inches(0.5), Inches(5.05), Inches(9.0), Inches(0.4),
                "実行フローとフィードバック", font_size=18, bold=True, color=BRAND_BLUE)
    steps = [
        "①対象選択", "②確認モーダル", "③Sidekiqジョブ\n（100件超時）",
        "④進捗表示", "⑤完了バナー\n＋メール通知",
    ]
    sx = 0.5
    for i, t in enumerate(steps):
        add_rounded_rect(s, Inches(sx), Inches(5.55),
                         Inches(1.55), Inches(0.85),
                         BRAND_BLUE if i in (1, 4) else BG_LIGHT,
                         line=LINE_GRAY)
        add_textbox(s, Inches(sx), Inches(5.65), Inches(1.55), Inches(0.7),
                    t, font_size=11, bold=(i in (1,4)),
                    color=WHITE if i in (1,4) else TEXT_BLACK,
                    align=PP_ALIGN.CENTER)
        if i < 4:
            add_textbox(s, Inches(sx + 1.55), Inches(5.78),
                        Inches(0.3), Inches(0.3),
                        "▶", font_size=14, color=BRAND_BLUE, align=PP_ALIGN.CENTER)
        sx += 1.85

    # 注釈
    add_textbox(s, Inches(0.5), Inches(6.7), Inches(9.0), Inches(0.3),
                "※100件超は非同期処理。完了後にメール通知 + 非同期タスク一覧から進捗確認可。",
                font_size=11, color=TEXT_GRAY)


# ===== Slide 8: 自動返信ルール UI（単店舗UIベース） =====
def slide_auto_reply_ui():
    s = prs.slides.add_slide(BLANK)
    add_master_decor(s, 8)
    add_slide_title(s, "5. 【画面B】自動返信ルール設定 UI設計")

    add_textbox(s, Inches(0.5), Inches(0.95), Inches(9.0), Inches(0.4),
                "単店舗の「クチコミ返信パターン・自動返信」画面のUIを踏襲し、グループ向けに「適用店舗」を追加",
                font_size=13, color=TEXT_GRAY)

    # ワイヤーフレーム
    wf_x, wf_y, wf_w, wf_h = 0.5, 1.45, 9.0, 5.45
    add_rect(s, Inches(wf_x), Inches(wf_y), Inches(wf_w), Inches(wf_h),
             WHITE, line=BRAND_BLUE, line_w=1.2)
    # ヘッダー
    add_rect(s, Inches(wf_x), Inches(wf_y), Inches(wf_w), Inches(0.40), BRAND_BLUE)
    add_textbox(s, Inches(wf_x + 0.15), Inches(wf_y + 0.05),
                Inches(8.0), Inches(0.32),
                "クチコミ管理 > 自動返信ルール設定", font_size=12, bold=True, color=WHITE)
    # 自動返信一括設定 右上
    add_textbox(s, Inches(wf_x + wf_w - 2.0), Inches(wf_y + 0.05),
                Inches(1.85), Inches(0.32),
                "☑ 自動返信 一括有効化",
                font_size=11, color=WHITE, align=PP_ALIGN.RIGHT)

    # 説明テキスト
    add_textbox(s, Inches(wf_x + 0.25), Inches(wf_y + 0.5),
                Inches(wf_w - 0.5), Inches(0.55),
                "返信パターンを最大20件登録できます。条件（評価・キーワード・店舗）に一致する新着クチコミに\n"
                "対し、毎時0分に自動で返信します。優先度はドラッグ＆ドロップで変更可能です。",
                font_size=11, color=TEXT_BLACK, line_spacing=1.3)

    # ===== パターンカード =====
    card_y = wf_y + 1.20
    card_w = wf_w - 0.5
    card_h = 1.50

    # カード背景
    add_rect(s, Inches(wf_x + 0.25), Inches(card_y),
             Inches(card_w), Inches(card_h), WHITE, line=BRAND_BLUE, line_w=1.0)
    # ドラッグハンドル
    add_textbox(s, Inches(wf_x + 0.35), Inches(card_y + 0.06),
                Inches(0.3), Inches(0.4),
                "⋮⋮", font_size=18, color=TEXT_GRAY)
    add_textbox(s, Inches(wf_x + 0.65), Inches(card_y + 0.08),
                Inches(5.0), Inches(0.3),
                "返信パターン名（20文字以内）  必須",
                font_size=11, bold=True, color=TEXT_BLACK)
    # トグル
    add_rounded_rect(s, Inches(wf_x + card_w - 1.0), Inches(card_y + 0.08),
                     Inches(0.75), Inches(0.28), BRAND_BLUE)
    add_textbox(s, Inches(wf_x + card_w - 1.0), Inches(card_y + 0.10),
                Inches(0.75), Inches(0.24),
                "● 有効", font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # 入力テキスト
    add_rect(s, Inches(wf_x + 0.65), Inches(card_y + 0.40),
             Inches(card_w - 1.5), Inches(0.28), BG_LIGHT, line=LINE_GRAY)
    add_textbox(s, Inches(wf_x + 0.72), Inches(card_y + 0.42),
                Inches(card_w - 1.6), Inches(0.24),
                "例：★1-2 低評価へのお詫び返信",
                font_size=10, color=TEXT_GRAY)

    # 条件行（左カラム） / 返信内容（右カラム）
    sub_y = card_y + 0.78
    # 左カラム見出し
    add_textbox(s, Inches(wf_x + 0.65), Inches(sub_y),
                Inches(4.0), Inches(0.28),
                "クチコミの内容（条件） 必須",
                font_size=11, bold=True, color=TEXT_BLACK)
    # 3つの条件ドロップダウン
    drops = ["感情分析評価 [▼]", "★数の選択 [▼]", "コメントの有無 [▼]"]
    for i, t in enumerate(drops):
        x = wf_x + 0.65 + (i % 3) * 1.45
        y = sub_y + 0.32
        add_rect(s, Inches(x), Inches(y), Inches(1.40), Inches(0.28),
                 WHITE, line=LINE_GRAY)
        add_textbox(s, Inches(x + 0.05), Inches(y + 0.02),
                    Inches(1.3), Inches(0.24),
                    t, font_size=9, color=TEXT_BLACK)
    # 適用店舗（新規追加）
    add_rounded_rect(s, Inches(wf_x + 0.65), Inches(sub_y + 0.68),
                     Inches(0.5), Inches(0.24), ACCENT_ORANGE)
    add_textbox(s, Inches(wf_x + 0.65), Inches(sub_y + 0.68),
                Inches(0.5), Inches(0.24),
                "NEW", font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(s, Inches(wf_x + 1.20), Inches(sub_y + 0.69),
                Inches(3.0), Inches(0.24),
                "適用店舗：● 全店舗(32)  ○ 個別選択",
                font_size=10, bold=True, color=TEXT_BLACK)

    # 右カラム
    rx = wf_x + 5.30
    add_textbox(s, Inches(rx), Inches(sub_y),
                Inches(3.5), Inches(0.28),
                "クチコミ返信内容 必須",
                font_size=11, bold=True, color=TEXT_BLACK)
    add_textbox(s, Inches(rx + 2.6), Inches(sub_y),
                Inches(1.3), Inches(0.28),
                "☑ 自動返信",
                font_size=10, color=TEXT_BLACK, align=PP_ALIGN.RIGHT)
    add_rect(s, Inches(rx), Inches(sub_y + 0.32),
             Inches(3.7), Inches(0.55), BG_LIGHT, line=LINE_GRAY)
    add_textbox(s, Inches(rx + 0.05), Inches(sub_y + 0.34),
                Inches(3.6), Inches(0.5),
                "この度はご来店いただきありがとうございました。\n貴重なご意見、誠にありがとうございます…",
                font_size=9, color=TEXT_GRAY)

    # 追加ボタン
    add_y = card_y + card_h + 0.20
    add_rounded_rect(s, Inches(wf_x + wf_w - 1.5), Inches(add_y),
                     Inches(1.25), Inches(0.32), WHITE, line=BRAND_BLUE, line_w=1.2)
    add_textbox(s, Inches(wf_x + wf_w - 1.5), Inches(add_y + 0.03),
                Inches(1.25), Inches(0.28),
                "+ パターン追加",
                font_size=10, bold=True, color=BRAND_BLUE, align=PP_ALIGN.CENTER)

    # 下部 操作ボタン
    btm_y = wf_y + wf_h - 0.55
    add_rounded_rect(s, Inches(wf_x + wf_w - 3.4), Inches(btm_y),
                     Inches(1.4), Inches(0.36), WHITE, line=LINE_GRAY)
    add_textbox(s, Inches(wf_x + wf_w - 3.4), Inches(btm_y + 0.05),
                Inches(1.4), Inches(0.28),
                "キャンセル", font_size=11, color=TEXT_BLACK, align=PP_ALIGN.CENTER)
    add_rounded_rect(s, Inches(wf_x + wf_w - 1.9), Inches(btm_y),
                     Inches(1.65), Inches(0.36), BRAND_BLUE)
    add_textbox(s, Inches(wf_x + wf_w - 1.9), Inches(btm_y + 0.05),
                Inches(1.65), Inches(0.28),
                "登録する", font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ===== Slide 9: 自動返信 機能詳細（単店舗との差分） =====
def slide_auto_reply_detail():
    s = prs.slides.add_slide(BLANK)
    add_master_decor(s, 9)
    add_slide_title(s, "5. 【画面B】自動返信ルール 機能詳細")

    # 単店舗からの拡張ポイント
    add_textbox(s, Inches(0.5), Inches(1.0), Inches(9.0), Inches(0.4),
                "単店舗UIからの拡張ポイント", font_size=18, bold=True, color=BRAND_BLUE)

    diffs = [
        ("適用店舗の追加", "「全店舗 / 個別選択」を新設。個別選択時は検索可能なチェックリストUI。"),
        ("優先度の概念", "上から順に評価。ドラッグ&ドロップで並び替え可能（⋮⋮ハンドル）。"),
        ("キーワード条件", "感情分析・★数に加え、本文中のキーワード「含む/含まない」を追加。"),
        ("1日の返信上限", "過度な自動返信を防止する1日あたりの上限件数設定（無制限可）。"),
        ("適用タイミング", "投稿後 即時 / N時間後（0-24）の遅延設定でクッション期間を確保。"),
    ]
    for i, (h, b) in enumerate(diffs):
        y = 1.50 + i * 0.62
        add_rounded_rect(s, Inches(0.5), Inches(y),
                         Inches(0.5), Inches(0.5), ACCENT_ORANGE)
        add_textbox(s, Inches(0.5), Inches(y + 0.07),
                    Inches(0.5), Inches(0.36),
                    str(i+1),
                    font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, Inches(1.15), Inches(y), Inches(2.5), Inches(0.5),
                    h, font_size=14, bold=True, color=TEXT_BLACK,
                    anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(s, Inches(3.7), Inches(y), Inches(5.8), Inches(0.5),
                    b, font_size=12, color=TEXT_BLACK, anchor=MSO_ANCHOR.MIDDLE)

    # 競合時の処理
    add_textbox(s, Inches(0.5), Inches(4.85), Inches(9.0), Inches(0.4),
                "ルール競合時の挙動", font_size=18, bold=True, color=BRAND_BLUE)
    add_rect(s, Inches(0.5), Inches(5.30),
             Inches(9.0), Inches(1.55), BG_LIGHT)
    bullet_text = (
        "• 複数ルールが同一クチコミにマッチした場合、優先度の高い（上位の）ルールのみを適用\n"
        "• 優先度の低いマッチルールはスキップ（ログに記録）\n"
        "• 「返信済みクチコミ」はデフォルトでスキップ。上書き有効化も選択可能\n"
        "• 1日の返信上限に達したルールは翌日まで待機"
    )
    add_textbox(s, Inches(0.7), Inches(5.45),
                Inches(8.6), Inches(1.4),
                bullet_text, font_size=13, color=TEXT_BLACK, line_spacing=1.4)


# ===== Slide 10: テンプレ管理 =====
def slide_template():
    s = prs.slides.add_slide(BLANK)
    add_master_decor(s, 10)
    add_slide_title(s, "6. 返信テンプレート管理")

    add_textbox(s, Inches(0.5), Inches(1.0), Inches(9.0), Inches(0.4),
                "一括返信・自動返信から共通参照する返信文テンプレート",
                font_size=14, color=TEXT_GRAY)

    # ワイヤーフレーム（簡略版）
    wf_x, wf_y, wf_w, wf_h = 0.5, 1.55, 9.0, 3.45
    add_rect(s, Inches(wf_x), Inches(wf_y), Inches(wf_w), Inches(wf_h),
             WHITE, line=BRAND_BLUE, line_w=1.2)
    add_rect(s, Inches(wf_x), Inches(wf_y), Inches(wf_w), Inches(0.40), BRAND_BLUE)
    add_textbox(s, Inches(wf_x + 0.15), Inches(wf_y + 0.05),
                Inches(7.0), Inches(0.32),
                "返信テンプレート管理", font_size=12, bold=True, color=WHITE)
    add_rounded_rect(s, Inches(wf_x + wf_w - 1.6), Inches(wf_y + 0.06),
                     Inches(1.45), Inches(0.28), WHITE)
    add_textbox(s, Inches(wf_x + wf_w - 1.6), Inches(wf_y + 0.08),
                Inches(1.45), Inches(0.24),
                "+ 新規追加",
                font_size=10, bold=True, color=BRAND_BLUE, align=PP_ALIGN.CENTER)

    # テーブル
    cols = ["テンプレート名", "本文プレビュー", "メモ", "操作"]
    col_w = [2.0, 4.0, 1.5, 1.3]
    tx = wf_x + 0.2
    ty = wf_y + 0.6
    add_rect(s, Inches(tx), Inches(ty), Inches(sum(col_w)), Inches(0.34), BG_HEAD)
    cx = tx
    for i, c in enumerate(cols):
        add_textbox(s, Inches(cx + 0.08), Inches(ty + 0.05),
                    Inches(col_w[i] - 0.16), Inches(0.24),
                    c, font_size=11, bold=True, color=TEXT_BLACK)
        cx += col_w[i]
    sample = [
        ("高評価への感謝", "この度は素晴らしいご評価をいただき誠にありがとうございます…",
         "★4-5用", "編集 削除"),
        ("低評価へのお詫び", "ご不便をおかけし誠に申し訳ございません。今後の改善に活かして…",
         "★1-2用", "編集 削除"),
        ("駐車場ご案内", "駐車場は店舗裏の○○駐車場をご利用いただけます…",
         "問い合わせ用", "編集 削除"),
        ("待ち時間お詫び", "お待たせしてしまい申し訳ございませんでした…",
         "クレーム用", "編集 削除"),
    ]
    for r_i, row in enumerate(sample):
        ry = ty + 0.34 + r_i * 0.50
        bg = WHITE if r_i % 2 == 0 else BG_LIGHT
        add_rect(s, Inches(tx), Inches(ry), Inches(sum(col_w)), Inches(0.50),
                 bg, line=LINE_GRAY)
        cx = tx
        for i, c in enumerate(row):
            color = BRAND_BLUE if i == 3 else TEXT_BLACK
            add_textbox(s, Inches(cx + 0.08), Inches(ry + 0.10),
                        Inches(col_w[i] - 0.16), Inches(0.34),
                        c, font_size=10, color=color)
            cx += col_w[i]

    # テンプレート項目
    add_textbox(s, Inches(0.5), Inches(5.20), Inches(9.0), Inches(0.4),
                "テンプレート設定項目", font_size=18, bold=True, color=BRAND_BLUE)
    head = ["項目", "型", "制約", "備考"]
    hx = [0.5, 2.4, 3.9, 6.4]
    hw = [1.85, 1.45, 2.45, 3.1]
    hy = 5.65
    for i, h in enumerate(head):
        add_rect(s, Inches(hx[i]), Inches(hy), Inches(hw[i]), Inches(0.32), BRAND_BLUE)
        add_textbox(s, Inches(hx[i]), Inches(hy + 0.03),
                    Inches(hw[i]), Inches(0.28),
                    h, font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    rows = [
        ("テンプレート名", "string", "必須・最大50文字", "一覧で表示される識別名"),
        ("本文", "text", "必須・最大300文字", "Googleの返信文字数上限に準拠"),
        ("メモ", "string", "任意・最大100文字", "内部運用メモ（返信文には含まれない）"),
    ]
    for r_i, row in enumerate(rows):
        y = hy + 0.32 + r_i * 0.36
        bg = WHITE if r_i % 2 == 0 else BG_LIGHT
        for j, c in enumerate(row):
            add_rect(s, Inches(hx[j]), Inches(y), Inches(hw[j]), Inches(0.36),
                     bg, line=LINE_GRAY)
            add_textbox(s, Inches(hx[j] + 0.08), Inches(y + 0.06),
                        Inches(hw[j] - 0.16), Inches(0.28),
                        c, font_size=10, color=TEXT_BLACK)


# ===== Slide 11: 設定項目・データモデル =====
def slide_data_model():
    s = prs.slides.add_slide(BLANK)
    add_master_decor(s, 11)
    add_slide_title(s, "7. 設定項目・データモデル")

    # 自動返信ルールの設定項目
    add_textbox(s, Inches(0.5), Inches(0.95), Inches(9.0), Inches(0.4),
                "自動返信ルール 設定項目一覧", font_size=18, bold=True, color=BRAND_BLUE)

    head = ["項目", "必須", "型", "バリデーション"]
    hx = [0.5, 3.2, 4.1, 5.5]
    hw = [2.65, 0.85, 1.35, 3.95]
    hy = 1.40
    for i, h in enumerate(head):
        add_rect(s, Inches(hx[i]), Inches(hy), Inches(hw[i]), Inches(0.32), BRAND_BLUE)
        add_textbox(s, Inches(hx[i]), Inches(hy + 0.03),
                    Inches(hw[i]), Inches(0.28),
                    h, font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    rows = [
        ("ルール名",          "○", "string",  "最大50文字"),
        ("適用対象",          "○", "enum",    "全店舗 / 個別店舗"),
        ("選択店舗",          "△", "array",   "適用対象=個別の場合1件以上"),
        ("評価条件",          "△", "array",   "★1-5の複数選択（未指定=全評価）"),
        ("キーワード条件",    "△", "array",   "{word, match_type} 最大10個 / 各50文字"),
        ("返信済みの扱い",    "○", "bool",    "スキップ（推奨） / 上書き"),
        ("返信モード",        "○", "enum",    "テンプレート / 固定文"),
        ("固定返信文",        "△", "text",    "返信モード=固定文の場合 / 最大300文字"),
        ("有効/無効",         "○", "bool",    "デフォルト: 有効"),
        ("適用タイミング",    "○", "int",     "0時間=即時 〜 24時間後"),
        ("1日の返信上限",     "△", "int",     "null=無制限 / 1〜9999"),
    ]
    for r_i, row in enumerate(rows):
        y = hy + 0.32 + r_i * 0.36
        bg = WHITE if r_i % 2 == 0 else BG_LIGHT
        for j, c in enumerate(row):
            add_rect(s, Inches(hx[j]), Inches(y), Inches(hw[j]), Inches(0.36),
                     bg, line=LINE_GRAY)
            color = ACCENT_ORANGE if (j == 1 and c == "○") else TEXT_BLACK
            bold = (j == 1 and c == "○")
            align = PP_ALIGN.CENTER if j == 1 else PP_ALIGN.LEFT
            add_textbox(s, Inches(hx[j] + 0.05), Inches(y + 0.06),
                        Inches(hw[j] - 0.1), Inches(0.28),
                        c, font_size=10, color=color, bold=bold, align=align)


# ===== Slide 12: 権限・実装フェーズ・要確認 =====
def slide_phase():
    s = prs.slides.add_slide(BLANK)
    add_master_decor(s, 12)
    add_slide_title(s, "8. 権限・実装フェーズ / 要確認事項")

    # 権限マトリクス
    add_textbox(s, Inches(0.5), Inches(0.95), Inches(9.0), Inches(0.4),
                "権限マトリクス", font_size=18, bold=True, color=BRAND_BLUE)
    head = ["操作", "master", "full", "edit", "view"]
    hx = [0.5, 3.8, 5.0, 6.2, 7.4]
    hw = [3.3, 1.2, 1.2, 1.2, 1.2]
    hy = 1.40
    for i, h in enumerate(head):
        add_rect(s, Inches(hx[i]), Inches(hy), Inches(hw[i]), Inches(0.30), BRAND_BLUE)
        add_textbox(s, Inches(hx[i]), Inches(hy + 0.02),
                    Inches(hw[i]), Inches(0.26),
                    h, font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    rows = [
        ("一括返信 実行",        "○", "○", "○", "✕"),
        ("自動返信ルール 作成",   "○", "○", "✕", "✕"),
        ("自動返信ルール 編集/削除", "○", "○", "✕", "✕"),
        ("自動返信ルール 参照",   "○", "○", "○", "○"),
        ("返信テンプレート 管理", "○", "○", "✕", "✕"),
    ]
    for r_i, row in enumerate(rows):
        y = hy + 0.30 + r_i * 0.30
        bg = WHITE if r_i % 2 == 0 else BG_LIGHT
        for j, c in enumerate(row):
            add_rect(s, Inches(hx[j]), Inches(y), Inches(hw[j]), Inches(0.30),
                     bg, line=LINE_GRAY)
            color = ACCENT_ORANGE if c == "✕" else (BRAND_BLUE if c == "○" else TEXT_BLACK)
            bold = (j > 0)
            align = PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT
            add_textbox(s, Inches(hx[j] + 0.05), Inches(y + 0.04),
                        Inches(hw[j] - 0.1), Inches(0.24),
                        c, font_size=10, color=color, bold=bold, align=align)

    # 実装フェーズ
    add_textbox(s, Inches(0.5), Inches(3.50), Inches(9.0), Inches(0.4),
                "実装フェーズ案", font_size=18, bold=True, color=BRAND_BLUE)
    phases = [
        ("Phase 1", "一括返信（固定文・テンプレ選択）", "M / 5-8日"),
        ("Phase 2", "自動返信ルール基本（評価条件）",   "M / 5-8日"),
        ("Phase 3", "キーワード条件・優先度並べ替え",   "S / 3-5日"),
        ("Phase 4", "AI返信文生成 連携",               "S / 3-5日"),
        ("Phase 5", "詳細ログ・レポート",              "S / 2-3日"),
    ]
    for i, (p, name, est) in enumerate(phases):
        y = 4.00 + i * 0.40
        add_rect(s, Inches(0.5), Inches(y), Inches(0.95), Inches(0.34), BRAND_BLUE)
        add_textbox(s, Inches(0.5), Inches(y + 0.05),
                    Inches(0.95), Inches(0.26),
                    p, font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_rect(s, Inches(1.45), Inches(y), Inches(6.8), Inches(0.34),
                 WHITE, line=LINE_GRAY)
        add_textbox(s, Inches(1.55), Inches(y + 0.05),
                    Inches(6.6), Inches(0.26),
                    name, font_size=11, color=TEXT_BLACK)
        add_rect(s, Inches(8.25), Inches(y), Inches(1.25), Inches(0.34),
                 BG_LIGHT, line=LINE_GRAY)
        add_textbox(s, Inches(8.25), Inches(y + 0.05),
                    Inches(1.25), Inches(0.26),
                    est, font_size=11, color=TEXT_GRAY, align=PP_ALIGN.CENTER)

    # 要確認事項
    add_textbox(s, Inches(0.5), Inches(6.10), Inches(9.0), Inches(0.4),
                "要確認事項", font_size=18, bold=True, color=BRAND_BLUE)
    add_rect(s, Inches(0.5), Inches(6.55), Inches(9.0), Inches(0.85),
             WARN_BG, line=ACCENT_ORANGE)
    add_textbox(s, Inches(0.7), Inches(6.62), Inches(8.7), Inches(0.78),
                "① Google Business Profile API の返信レート制限（上限/日）の最終値\n"
                "② 自動返信処理頻度（毎時 / 30分毎 / 即時Webhook）の方針\n"
                "③ 既存 pattern_auto_replies テーブルとの統合 or 並存",
                font_size=11, color=TEXT_BLACK, line_spacing=1.35)


# ===== Slide 13: 終了 =====
def slide_end():
    s = prs.slides.add_slide(BLANK)
    add_right_triangle(s, Inches(-1.0), Inches(0), Inches(6.5), Inches(7.5),
                       DEEP_NAVY, flip_h=True)
    add_right_triangle(s, Inches(7.0), Inches(-1.0), Inches(4.0), Inches(4.0),
                       BRAND_BLUE)
    add_right_triangle(s, Inches(-0.5), Inches(6.4), Inches(1.7), Inches(1.5),
                       BRAND_BLUE, flip_v=True)

    logo = s.shapes.add_textbox(Inches(3.4), Inches(3.3), Inches(3.5), Inches(0.8))
    tf = logo.text_frame
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r1 = p.add_run(); r1.text = "GMO"
    r1.font.name = "Noto Sans JP"; r1.font.size = Pt(40); r1.font.bold = True
    r1.font.color.rgb = BRAND_BLUE
    r2 = p.add_run(); r2.text = "TECH"
    r2.font.name = "Noto Sans JP"; r2.font.size = Pt(40); r2.font.bold = True
    r2.font.color.rgb = TEXT_GRAY

    add_textbox(s, Inches(2.0), Inches(4.3), Inches(6.0), Inches(0.4),
                "Thank you", font_size=20, color=TEXT_GRAY, align=PP_ALIGN.CENTER)

    add_textbox(s, Inches(4.5), Inches(7.1), Inches(5.5), Inches(0.3),
                "Copyright © GMO TECH All Rights reserved.   Confidential",
                font_size=10, color=TEXT_GRAY, align=PP_ALIGN.RIGHT)


# ===== 生成 =====
slide_title()
slide_toc()
slide_purpose()
slide_gap()
slide_structure()
slide_bulk_reply_ui()
slide_bulk_reply_detail()
slide_auto_reply_ui()
slide_auto_reply_detail()
slide_template()
slide_data_model()
slide_phase()
slide_end()

out = "/home/user/mixdasyuboardfuji/docs/group_review_reply_settings_spec.pptx"
prs.save(out)
print(f"saved: {out}")
print(f"slides: {len(prs.slides)}")
