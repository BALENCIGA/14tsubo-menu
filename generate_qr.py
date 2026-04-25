"""デザイン入りQRコード生成スクリプト。

イッテンヨンツボのメニューサイト用に、ブランドカラー（黒背景＋アードベッググリーン）の
QRコードを生成する。中央にロゴ、下部に「SCAN FOR MENU / @1.4tsubo」キャプションを配置。

使い方:
    python3 generate_qr.py

サイズや出力先、URLは MAIN 内のリストで指定する。
"""
from pathlib import Path

import qrcode
from PIL import Image, ImageDraw, ImageFont
from qrcode.constants import ERROR_CORRECT_H

# --- ブランドカラー（index.html の CSS変数から） ---
BG = (5, 8, 5)            # --bg #050805
NEON = (61, 245, 138)     # --neon #3df58a
INK_DIM = (143, 167, 144) # --ink-dim #8fa790

LOGO_PATH = Path(__file__).parent / "assets" / "logo" / "logo.png"

# ベースURL（QR/Instagramそれぞれにutmを付与）
BASE_URL = "https://balenciga.github.io/14tsubo-menu/"


def make_qr_image(url: str, size_px: int) -> Image.Image:
    """指定URL/サイズのQRコード画像（背景透過、QRが緑）を返す。"""
    qr = qrcode.QRCode(
        version=None,
        error_correction=ERROR_CORRECT_H,  # ロゴ重畳のため高誤り訂正
        box_size=10,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color=NEON, back_color=BG).convert("RGBA")
    img = img.resize((size_px, size_px), Image.LANCZOS)
    return img


def overlay_logo(qr_img: Image.Image, logo_ratio: float = 0.22) -> Image.Image:
    """QR画像の中央にロゴを重ねる。"""
    if not LOGO_PATH.exists():
        return qr_img
    logo = Image.open(LOGO_PATH).convert("RGBA")
    w = int(qr_img.size[0] * logo_ratio)
    aspect = logo.size[1] / logo.size[0]
    logo = logo.resize((w, int(w * aspect)), Image.LANCZOS)

    # ロゴ背後に黒い四角を敷いてQRが透けないようにする
    pad = 8
    bg_box = Image.new("RGBA", (logo.size[0] + pad * 2, logo.size[1] + pad * 2), BG + (255,))
    cx = (qr_img.size[0] - bg_box.size[0]) // 2
    cy = (qr_img.size[1] - bg_box.size[1]) // 2
    qr_img.alpha_composite(bg_box, (cx, cy))
    qr_img.alpha_composite(logo, (cx + pad, cy + pad))
    return qr_img


def compose_card(qr_img: Image.Image, canvas_size: tuple[int, int], qr_top: int, caption_y: int, caption_size: int) -> Image.Image:
    """黒背景キャンバスにQRとキャプションを配置した最終画像を返す。"""
    canvas = Image.new("RGBA", canvas_size, BG + (255,))
    qr_x = (canvas.size[0] - qr_img.size[0]) // 2
    canvas.alpha_composite(qr_img, (qr_x, qr_top))

    draw = ImageDraw.Draw(canvas)
    try:
        # システムフォント。Helvetica系で揃える
        font_main = ImageFont.truetype("/System/Library/Fonts/HelveticaNeue.ttc", caption_size)
        font_sub = ImageFont.truetype("/System/Library/Fonts/HelveticaNeue.ttc", int(caption_size * 0.55))
    except OSError:
        font_main = ImageFont.load_default()
        font_sub = ImageFont.load_default()

    main_text = "SCAN FOR MENU"
    sub_text = "@1.4tsubo"

    bbox = draw.textbbox((0, 0), main_text, font=font_main)
    tx = (canvas.size[0] - (bbox[2] - bbox[0])) // 2
    draw.text((tx, caption_y), main_text, fill=NEON, font=font_main)

    sub_y = caption_y + int(caption_size * 1.4)
    bbox2 = draw.textbbox((0, 0), sub_text, font=font_sub)
    tx2 = (canvas.size[0] - (bbox2[2] - bbox2[0])) // 2
    draw.text((tx2, sub_y), sub_text, fill=INK_DIM, font=font_sub)

    return canvas


def build(url: str, out_path: Path, canvas_size: tuple[int, int], qr_size: int, qr_top: int, caption_y: int, caption_size: int) -> None:
    """1枚のQR画像を生成して保存する。"""
    qr_img = make_qr_image(url, qr_size)
    qr_img = overlay_logo(qr_img)
    card = compose_card(qr_img, canvas_size, qr_top, caption_y, caption_size)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    card.convert("RGB").save(out_path, "PNG", optimize=True)
    print(f"✓ {out_path.name}: {url}")


def main() -> None:
    """QRサイズ × URL の組み合わせ分だけ生成する。"""
    out_dir = Path(__file__).parent / "pdf"

    # utm付きURL（GA4で流入元を区別するため）
    qr_url = f"{BASE_URL}?utm_source=qr&utm_medium=print&utm_campaign=menu"

    # 小サイズ（カード/ステッカー用、900x980）
    build(
        qr_url,
        out_dir / "14tsubo_qr.png",
        canvas_size=(900, 980),
        qr_size=820,
        qr_top=40,
        caption_y=870,
        caption_size=42,
    )

    # B5サイズ（卓上ポップ印刷用、2150x3035）
    build(
        qr_url,
        out_dir / "14tsubo_qr_b5.png",
        canvas_size=(2150, 3035),
        qr_size=1100,
        qr_top=900,
        caption_y=2100,
        caption_size=80,
    )


if __name__ == "__main__":
    main()
