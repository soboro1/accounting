from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


OUT = Path("/Users/ikenoharuhiko/Desktop/　学習")
FONT_DIR = Path("/Users/ikenoharuhiko/Library/Fonts")

REGULAR = FONT_DIR / "NotoSansCJKjp-Regular.otf"
MEDIUM = FONT_DIR / "NotoSansCJKjp-Medium.otf"
BOLD = FONT_DIR / "NotoSansCJKjp-Bold.otf"
BLACK = FONT_DIR / "NotoSansCJKjp-Black.otf"

BG = "#EEF1F5"
CARD = "#FFFFFF"
SOFT = "#F4F6F9"
NAVY = "#14233D"
GRAY = "#657187"
GOLD = "#A87310"
GOLD_DARK = "#8F610A"
BEIGE = "#F6F0E5"
BORDER = "#DCC79D"
LIGHT_BORDER = "#D7DDE7"
RED = "#C94B22"
GREEN = "#3F7A2C"


def f(size, weight="regular"):
    paths = {"regular": REGULAR, "medium": MEDIUM, "bold": BOLD, "black": BLACK}
    return ImageFont.truetype(str(paths[weight]), size)


def rr(draw, box, radius=24, fill=CARD, outline=None, width=2):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def text_width(draw, text, font):
    b = draw.textbbox((0, 0), text, font=font)
    return b[2] - b[0]


def wrap_text(draw, text, font, max_width):
    lines = []
    for paragraph in text.split("\n"):
        if not paragraph:
            lines.append("")
            continue
        current = ""
        for ch in paragraph:
            trial = current + ch
            if current and text_width(draw, trial, font) > max_width:
                lines.append(current)
                current = ch
            else:
                current = trial
        if current:
            lines.append(current)
    return lines


def draw_wrapped(draw, text, x, y, font, fill, max_width, line_gap=12, anchor=None):
    lines = wrap_text(draw, text, font, max_width)
    bbox = draw.textbbox((0, 0), "あA", font=font)
    line_h = bbox[3] - bbox[1]
    for line in lines:
        draw.text((x, y), line, font=font, fill=fill, anchor=anchor)
        y += line_h + line_gap
    return y


def pill(draw, xy, text, font, fill=BEIGE, outline=BORDER, text_fill=GOLD_DARK, pad_x=28, h=58):
    x, y = xy
    w = text_width(draw, text, font) + pad_x * 2
    rr(draw, (x, y, x + w, y + h), radius=h // 2, fill=fill, outline=outline, width=2)
    draw.text((x + w / 2, y + h / 2 - 1), text, font=font, fill=text_fill, anchor="mm")
    return w


def centered_text(draw, box, text, font, fill=NAVY):
    x1, y1, x2, y2 = box
    draw.text(((x1 + x2) / 2, (y1 + y2) / 2 - 1), text, font=font, fill=fill, anchor="mm")


def draw_branch_icon(draw, cx, cy, scale=1.0):
    line = max(4, int(5 * scale))
    gold = GOLD
    # central office
    draw.rounded_rectangle((cx - 72*scale, cy - 46*scale, cx + 72*scale, cy + 46*scale),
                           radius=int(14*scale), outline=gold, width=line, fill="#FFFDFC")
    draw.text((cx, cy - 7*scale), "本店", font=f(int(28*scale), "bold"), fill=NAVY, anchor="mm")
    # connector
    draw.line((cx, cy + 48*scale, cx, cy + 87*scale), fill=gold, width=line)
    draw.line((cx - 150*scale, cy + 87*scale, cx + 150*scale, cy + 87*scale), fill=gold, width=line)
    for dx, label in [(-150, "A支店"), (150, "B支店")]:
        x = cx + dx*scale
        draw.line((x, cy + 87*scale, x, cy + 116*scale), fill=gold, width=line)
        draw.rounded_rectangle((x - 68*scale, cy + 116*scale, x + 68*scale, cy + 196*scale),
                               radius=int(14*scale), outline=gold, width=line, fill="#FFFDFC")
        draw.text((x, cy + 155*scale), label, font=f(int(24*scale), "bold"), fill=NAVY, anchor="mm")


def make_header():
    im = Image.new("RGB", (1280, 670), BG)
    d = ImageDraw.Draw(im)
    rr(d, (34, 34, 1246, 636), radius=38, fill=CARD, outline=BORDER, width=3)

    pill(d, (82, 78), "日商簿記1級｜本支店会計", f(25, "bold"), h=56)
    d.text((82, 176), "本支店会計の理論問題", font=f(54, "black"), fill=NAVY)
    d.text((84, 260), "制度・内部利益・総合損益を4問で整理", font=f(29, "medium"), fill=GRAY)

    rr(d, (82, 348, 665, 484), radius=24, fill=SOFT, outline=LIGHT_BORDER, width=2)
    d.text((116, 382), "数字ではなく、", font=f(24, "medium"), fill=GRAY)
    d.text((116, 425), "用語と考え方を固める", font=f(33, "bold"), fill=GOLD_DARK)

    # accent dots
    for x in (90, 116, 142):
        d.ellipse((x, 548, x + 11, 559), fill=GOLD)
    d.text((168, 552), "オリジナル練習問題｜全4問", font=f(22, "medium"), fill=GRAY, anchor="lm")
    draw_branch_icon(d, 955, 251, 1.0)
    im.save(OUT / "noteヘッダー_本支店会計理論.png", quality=95)


QUESTIONS = [
    (
        "問1",
        "本店が支店相互間の取引を一括して記帳し、各支店が相手勘定として「本店」を用いる制度はどれか。",
        ["本店集中計算制度", "支店分散計算制度", "独立会計制度", "総合損益方式"],
    ),
    (
        "問2",
        "支店分散計算制度で、A支店がB支店から商品を受け取ったとき、A支店が用いる相手勘定はどれか。",
        ["本店", "A支店", "B支店", "仕入"],
    ),
    (
        "問3",
        "本店から支店へ振替価格で発送した商品の一部が、支店の期末商品に含まれている。振替価格に含まれる利益を何というか。",
        ["売上総利益", "未実現利益（内部利益）", "為替差損益", "評価差額"],
    ),
    (
        "問4",
        "本店と各支店の損益を集計し、会社全体の当期純損益を計算するために本店が用いる勘定はどれか。",
        ["支店勘定", "本店勘定", "総合損益勘定", "繰越利益剰余金"],
    ),
]


def draw_question_block(d, y, label, q, options):
    rr(d, (58, y, 1542, y + 300), radius=28, fill=CARD, outline=BORDER, width=2)
    d.text((96, y + 39), label, font=f(28, "bold"), fill=GOLD_DARK)
    draw_wrapped(d, q, 190, y + 36, f(27, "medium"), NAVY, 1280, line_gap=8)
    option_y = y + 170
    start_x, gap, w, h = 96, 18, 338, 76
    for i, option in enumerate(options):
        x = start_x + i * (w + gap)
        rr(d, (x, option_y, x + w, option_y + h), radius=38, fill="#FFFEFC", outline=BORDER, width=2)
        fs = 24 if len(option) <= 10 else 21
        centered_text(d, (x, option_y, x + w, option_y + h), option, f(fs, "medium"), NAVY)


def make_problem():
    im = Image.new("RGB", (1600, 1660), BG)
    d = ImageDraw.Draw(im)
    rr(d, (24, 24, 1576, 1636), radius=34, fill="#F8F9FB", outline=BORDER, width=3)
    pill(d, (70, 62), "出題パターン：理論（本支店会計）", f(24, "bold"), h=56)
    d.text((70, 151), "本支店会計の用語と考え方", font=f(42, "black"), fill=NAVY)
    d.text((70, 218), "次の文章に最も適切な語句を、選択肢から選びなさい。", font=f(25, "medium"), fill=GRAY)
    for idx, (label, q, opts) in enumerate(QUESTIONS):
        draw_question_block(d, 286 + idx * 316, label, q, opts)
    d.text((70, 1588), "※ 同じ単元を別の聞き方で確認するオリジナル問題です。", font=f(19, "regular"), fill=GRAY)
    im.save(OUT / "本支店会計理論_問題.png", quality=95)


ANSWERS = [
    (
        "問1",
        "本店集中計算制度",
        "本店集中計算制度では、支店どうしの取引も本店を経由して記帳します。支店側の相手勘定は「本店」、本店側は各支店勘定で整理します。",
    ),
    (
        "問2",
        "B支店",
        "支店分散計算制度では、各支店が相手支店の勘定を直接使います。A支店は、商品を送ったB支店の勘定を用います。",
    ),
    (
        "問3",
        "未実現利益（内部利益）",
        "振替価格に含まれる利益のうち、外部販売されていない期末商品分は会社全体では未実現です。合併上は商品から控除して原価に戻します。",
    ),
    (
        "問4",
        "総合損益勘定",
        "各支店の損益を本店の「総合損益」勘定へ集め、会社全体の当期純損益を確定します。確定後は繰越利益剰余金へ振り替えます。",
    ),
]


def draw_answer_card(d, y, label, answer, explanation):
    rr(d, (58, y, 1542, y + 252), radius=28, fill=CARD, outline=LIGHT_BORDER, width=2)
    d.ellipse((92, y + 39, 142, y + 89), fill=GREEN)
    d.text((117, y + 63), "✓", font=f(28, "bold"), fill="white", anchor="mm")
    d.text((164, y + 37), f"{label}  正解", font=f(25, "bold"), fill=GRAY)
    d.text((164, y + 82), answer, font=f(32, "bold"), fill=GOLD_DARK)
    draw_wrapped(d, explanation, 164, y + 142, f(24, "regular"), NAVY, 1295, line_gap=8)


def make_answer():
    im = Image.new("RGB", (1600, 1660), BG)
    d = ImageDraw.Draw(im)
    rr(d, (24, 24, 1576, 1636), radius=34, fill="#F8F9FB", outline=BORDER, width=3)
    d.text((70, 68), "解答・解説", font=f(45, "black"), fill=NAVY)
    d.text((70, 135), "正解だけでなく、制度ごとの『相手勘定』まで確認しましょう。", font=f(24, "medium"), fill=GRAY)

    for idx, item in enumerate(ANSWERS):
        draw_answer_card(d, 205 + idx * 270, *item)

    y = 1302
    rr(d, (58, y, 1542, 1576), radius=28, fill=SOFT, outline=LIGHT_BORDER, width=2)
    d.text((94, y + 38), "4問のまとめ", font=f(29, "bold"), fill=GOLD_DARK)
    summaries = [
        "本店集中：支店は「本店」勘定を使う",
        "支店分散：相手支店の勘定を直接使う",
        "期末内部利益：未実現なので商品から控除",
        "全社損益：総合損益勘定でまとめる",
    ]
    positions = [(94, y + 103), (802, y + 103), (94, y + 181), (802, y + 181)]
    for i, (summary, (x, yy)) in enumerate(zip(summaries, positions), 1):
        d.ellipse((x, yy, x + 45, yy + 45), fill=BEIGE, outline=BORDER, width=2)
        d.text((x + 22.5, yy + 21), str(i), font=f(21, "bold"), fill=GOLD_DARK, anchor="mm")
        d.text((x + 60, yy + 21), summary, font=f(21, "medium"), fill=NAVY, anchor="lm")
    d.text((70, 1600), "日商簿記1級・本支店会計の復習用（オリジナル問題）", font=f(18, "regular"), fill=GRAY)
    im.save(OUT / "本支店会計理論_解答解説.png", quality=95)


if __name__ == "__main__":
    make_header()
    make_problem()
    make_answer()
    print("created:")
    for name in [
        "noteヘッダー_本支店会計理論.png",
        "本支店会計理論_問題.png",
        "本支店会計理論_解答解説.png",
    ]:
        print(OUT / name)
