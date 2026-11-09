# -*- coding: utf-8 -*-
"""界面主题：从背景图与图标提取配色，并生成可读的模糊背景。

设计原则：
- 主色来自 background.png / icon.png 的主色调（墨绿 + 薄荷绿 + 暖白）；
- 面板使用同色系深色，保证文字对比度；
- 若提取失败或颜色不适合，回退到内置的安全配色。
"""
from __future__ import annotations

import colorsys
import random
from collections import OrderedDict
from dataclasses import dataclass
from typing import Dict, List, Tuple

from PIL import Image, ImageChops, ImageDraw, ImageEnhance, ImageFilter

RGB = Tuple[int, int, int]

_CHARACTER_MASTER_CACHE: Dict[str, Image.Image] = {}
_CHARACTER_SCALED_CACHE: "OrderedDict[Tuple[str, int, int, int], Image.Image]" = OrderedDict()
_CHARACTER_CACHE_LIMIT = 8

AMBIENT_SEED = 20261104
AMBIENT_BASE_SIZE = (1180, 740)
_AMBIENT_CACHE: "OrderedDict[Tuple[int, int], Image.Image]" = OrderedDict()
_AMBIENT_CACHE_LIMIT = 3
_AMBIENT_LAYOUT = None
_AMBIENT_COLOR_CACHE: Dict[str, Tuple[RGB, RGB, RGB]] = {}


def _rgb(value) -> RGB:
    return (int(value[0]), int(value[1]), int(value[2]))


def _hex(value: RGB) -> str:
    return "#%02X%02X%02X" % (int(value[0]), int(value[1]), int(value[2]))


def _luminance(color: RGB) -> float:
    r, g, b = color
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0


def _saturation(color: RGB) -> float:
    r, g, b = [v / 255.0 for v in color]
    h, l, s = colorsys.rgb_to_hls(r, g, b)[0], colorsys.rgb_to_hls(r, g, b)[1], colorsys.rgb_to_hls(r, g, b)[2]
    return s


def _mix(a: RGB, b: RGB, t: float) -> RGB:
    return (
        int(a[0] + (b[0] - a[0]) * t),
        int(a[1] + (b[1] - a[1]) * t),
        int(a[2] + (b[2] - a[2]) * t),
    )


def _set_luminance(color: RGB, target: float) -> RGB:
    """在保持色相的前提下拉高/压低亮度。"""
    h, l, s = colorsys.rgb_to_hls(color[0] / 255.0, color[1] / 255.0, color[2] / 255.0)
    l = max(0.0, min(1.0, target))
    return tuple(int(round(v * 255)) for v in colorsys.hls_to_rgb(h, l, s))


def _is_green(color: RGB) -> bool:
    r, g, b = color
    return g >= r and g >= b and (g - min(r, b)) >= 8


def _contrast_text(background: RGB) -> str:
    return "#F7FBF6" if _luminance(background) < 0.55 else "#142018"


@dataclass
class Palette:
    accent: str
    accent_dark: str
    accent_soft: str
    panel: str
    panel_alt: str
    panel_raised: str
    text: str
    text_muted: str
    border: str
    success: str
    warning: str
    error: str
    log_bg: str
    log_text: str


def extract_colors(image_path: str, limit: int = 14) -> List[RGB]:
    """提取主色：先合成到白底（处理透明背景），再量化统计。"""
    with Image.open(image_path) as src:
        im = src.convert("RGBA")
        size = im.size
        if size[0] > 240 or size[1] > 240:
            ratio = min(240.0 / max(size[0], 1), 240.0 / max(size[1], 1))
            im = im.resize((max(1, int(size[0] * ratio)), max(1, int(size[1] * ratio))), Image.LANCZOS)
        background = Image.new("RGBA", im.size, (255, 255, 255, 255))
        background.alpha_composite(im)
        rgb = background.convert("RGB")
    quantized = rgb.quantize(colors=limit, method=Image.MEDIANCUT)
    palette = quantized.getpalette() or []
    colors = []
    for count, index in sorted(quantized.getcolors() or [], reverse=True):
        r, g, b = palette[index * 3:index * 3 + 3]
        colors.append((count, (r, g, b)))
    return colors


def build_palette(background_path: str, icon_path: str = "") -> Palette:
    colors = []
    for path in (background_path, icon_path):
        if not path:
            continue
        try:
            colors.extend(extract_colors(path))
        except Exception:
            continue
    # 过滤掉接近纯白/纯黑/大面积透明底的色块，优先绿色系，其次任意饱和色。
    candidates = []
    for count, color in colors:
        if count < 20:
            continue
        lum = _luminance(color)
        if lum > 0.93 or lum < 0.06:
            continue
        if _saturation(color) < 0.06:
            continue
        candidates.append(color)
    greens = [c for c in candidates if _is_green(c)]
    base = greens[0] if greens else (candidates[0] if candidates else (120, 160, 120))

    # 用提取到的绿色生成深色面板；亮度刻意压低，确保与浅色文字对比。
    panel = _set_luminance(_mix(base, (10, 24, 18), 0.72), 0.105)
    panel_alt = _set_luminance(_mix(base, (10, 24, 18), 0.62), 0.145)
    panel_raised = _set_luminance(_mix(base, (10, 24, 18), 0.52), 0.185)
    accent = _set_luminance(base, 0.68)
    accent_dark = _set_luminance(base, 0.40)
    accent_soft = _mix(accent, (255, 255, 255), 0.45)

    text = "#F7FBF4"
    text_muted = _hex(_mix(accent_soft, (230, 240, 230), 0.35))
    border = _hex(_mix(accent, panel, 0.55))
    return Palette(
        accent=_hex(accent),
        accent_dark=_hex(accent_dark),
        accent_soft=_hex(accent_soft),
        panel=_hex(panel),
        panel_alt=_hex(panel_alt),
        panel_raised=_hex(panel_raised),
        text=text,
        text_muted=text_muted,
        border=border,
        success=_hex(_set_luminance((110, 200, 140), 0.68)),
        warning=_hex(_set_luminance((238, 194, 110), 0.70)),
        error=_hex(_set_luminance((238, 120, 110), 0.66)),
        log_bg="#0C1611",
        log_text="#DCEBE0",
    )


def _ambient_colors(background_path: str) -> Tuple[RGB, RGB, RGB]:
    """从背景图主色聚类中挑出深绿、浅绿，并由浅绿偏移出协调的蓝。

    返回 (深绿, 浅绿, 蓝)。蓝不是纯饱和原色：色相转到蓝相、饱和度压到 0.30 左右。
    """
    cached = _AMBIENT_COLOR_CACHE.get(background_path)
    if cached is not None:
        return cached
    greens: List[RGB] = []
    try:
        for _count, color in extract_colors(background_path):
            lum = _luminance(color)
            if _is_green(color) and 0.25 <= lum <= 0.9 and _saturation(color) >= 0.12:
                greens.append(color)
    except Exception:
        greens = []
    greens.sort(key=_luminance)
    deep = _set_luminance(greens[0], 0.50) if greens else (78, 116, 88)
    light = _set_luminance(greens[-1], 0.84) if greens else (180, 216, 174)
    blue = tuple(int(round(v * 255)) for v in colorsys.hls_to_rgb(0.57, 0.66, 0.34))
    colors = (tuple(deep), tuple(light), tuple(blue))
    _AMBIENT_COLOR_CACHE[background_path] = colors
    return colors


def _cubic(p0, p1, p2, p3, steps: int = 26) -> List[Tuple[float, float]]:
    pts = []
    for i in range(steps + 1):
        t = i / float(steps)
        mt = 1.0 - t
        x = (mt ** 3) * p0[0] + 3 * (mt ** 2) * t * p1[0] + 3 * mt * (t ** 2) * p2[0] + (t ** 3) * p3[0]
        y = (mt ** 3) * p0[1] + 3 * (mt ** 2) * t * p1[1] + 3 * mt * (t ** 2) * p2[1] + (t ** 3) * p3[1]
        pts.append((x, y))
    return pts


def _petal_mask(width: int, height: int) -> Image.Image:
    """花瓣：尖端朝上，中脉镂空出一道细缝。"""
    ss = 3
    tw, th = width * ss, height * ss
    mask = Image.new("L", (tw, th), 0)
    draw = ImageDraw.Draw(mask)
    left = _cubic((tw * 0.5, th * 0.02), (tw * 0.15, th * 0.28),
                  (tw * 0.04, th * 0.70), (tw * 0.5, th * 0.98))
    right = _cubic((tw * 0.5, th * 0.02), (tw * 0.85, th * 0.28),
                   (tw * 0.96, th * 0.70), (tw * 0.5, th * 0.98))
    draw.polygon(left + list(reversed(right)), fill=255)
    draw.line([(tw * 0.5, th * 0.14), (tw * 0.5, th * 0.88)],
              fill=0, width=max(1, int(tw * 0.04)))
    return mask


def _computer_mask(width: int, height: int) -> Image.Image:
    """小电脑：显示器 + 镂空屏幕 + 底座。"""
    ss = 3
    tw, th = width * ss, height * ss
    mask = Image.new("L", (tw, th), 0)
    draw = ImageDraw.Draw(mask)
    body_h = int(th * 0.68)
    radius = int(tw * 0.08)
    draw.rounded_rectangle([0, 0, tw - 1, body_h], radius=radius, fill=255)
    inset = int(tw * 0.11)
    draw.rounded_rectangle(
        [inset, int(th * 0.09), tw - 1 - inset, body_h - int(th * 0.10)],
        radius=max(1, int(radius * 0.7)), fill=0,
    )
    draw.polygon([(tw * 0.40, body_h - 1), (tw * 0.60, body_h - 1),
                  (tw * 0.56, th * 0.88), (tw * 0.44, th * 0.88)], fill=255)
    draw.rounded_rectangle([tw * 0.20, th * 0.88, tw * 0.80, th - 1],
                           radius=max(1, int(tw * 0.05)), fill=255)
    draw.ellipse([tw * 0.46, body_h - int(th * 0.07), tw * 0.54,
                  body_h - int(th * 0.03)], fill=0)
    return mask


def _floppy_mask(width: int, height: int) -> Image.Image:
    """软盘：切角外壳、镂空标签、读写窗与滑块。"""
    ss = 3
    tw, th = width * ss, height * ss
    mask = Image.new("L", (tw, th), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([0, 0, tw - 1, th - 1],
                           radius=max(1, int(tw * 0.05)), fill=255)
    draw.polygon([(tw * 0.80, 0), (tw - 1, 0), (tw - 1, th * 0.20)], fill=0)
    draw.rounded_rectangle([tw * 0.16, th * 0.36, tw * 0.84, th * 0.82],
                           radius=max(1, int(tw * 0.035)), fill=0)
    draw.rectangle([tw * 0.40, th * 0.07, tw * 0.76, th * 0.24], fill=0)
    draw.rectangle([tw * 0.55, th * 0.09, tw * 0.74, th * 0.22], fill=255)
    return mask


def _shape_sprite(kind: str, size: int, color: RGB, angle: float) -> Image.Image:
    """生成单个旋转后的装饰图形，超采样后缩小获得平滑边缘。"""
    ss = 3
    if kind == "petal":
        width, height = int(size * 0.78), int(size)
        mask = _petal_mask(width, height)
    elif kind == "computer":
        width, height = int(size), int(size * 0.92)
        mask = _computer_mask(width, height)
    else:
        width, height = int(size * 0.92), int(size)
        mask = _floppy_mask(width, height)
    if angle:
        mask = mask.rotate(angle, resample=Image.BICUBIC, expand=True)
    mask = mask.resize((max(1, mask.width // ss), max(1, mask.height // ss)), Image.LANCZOS)
    sprite = Image.new("RGBA", mask.size, (color[0], color[1], color[2], 0))
    sprite.putalpha(mask)
    return sprite


def _ambient_layout():
    """固定种子生成一次布局（基准画布坐标），窗口缩放时按比例复用，图形不会跳位。

    密度收敛成“零星点缀”的 8 个（花瓣 3、小电脑 2、软盘 3）：页眉 2 个，
    其余落在四角与边距，锚点彼此拉开，避免堆叠或挤在同一块。
    """
    global _AMBIENT_LAYOUT
    if _AMBIENT_LAYOUT is not None:
        return _AMBIENT_LAYOUT
    rng = random.Random(AMBIENT_SEED)
    base_w, base_h = AMBIENT_BASE_SIZE
    # (kind, x, y)：页眉放 2 个（右侧亮区 1 个可辨认），其余分布在四角与边距。
    anchors = [
        ("floppy", 890, 96),      # 页眉中部亮区：唯一不被压暗的装饰
        ("petal", 600, 96),       # 页眉中部，避开左侧标题文字
        ("computer", 180, 250),   # 左上
        ("floppy", 1060, 330),    # 右上
        ("floppy", 430, 470),     # 中部偏左，只在面板里透出淡影
        ("petal", 150, 600),      # 左下
        ("petal", 1060, 600),     # 右下
        ("computer", 760, 594),   # 底部中间
    ]
    # 0 浅绿、1 深绿、2 蓝：只剩 8 个时按 3/3/2 预设再打乱，保证三种色调都出现。
    color_slots = [0, 0, 0, 1, 1, 1, 2, 2]
    rng.shuffle(color_slots)
    placed = []
    for index, (kind, cx, cy) in enumerate(anchors):
        if kind == "petal":
            size = rng.uniform(117, 279)
        elif kind == "computer":
            size = rng.uniform(180, 327)
        else:
            size = rng.uniform(144, 243)
        half_w = size * 0.45
        half_h = size * 0.50
        cx = min(max(cx, half_w + 8), base_w - half_w - 8)
        cy = min(max(cy, half_h + 6), base_h - half_h - 6)
        if kind == "petal":
            angle = rng.uniform(0, 360)
        elif kind == "computer":
            angle = rng.uniform(-12, 12)
        else:
            angle = rng.uniform(-15, 15)
        # 0 浅绿、1 深绿、2 蓝，深浅混合避免整片同色。
        color_index = color_slots[index]
        placed.append((kind, cx, cy, size, angle, color_index))
    _AMBIENT_LAYOUT = placed
    return placed


def _ambient_base(size: Tuple[int, int]) -> Image.Image:
    """深绿竖向渐变底，作为装饰图形的衬底。"""
    width, height = size
    top = Image.new("RGB", (width, height), (17, 36, 26))
    bottom = Image.new("RGB", (width, height), (9, 19, 13))
    ramp = Image.linear_gradient("L").resize((width, height), Image.BILINEAR)
    return Image.composite(bottom, top, ramp)


def ambient_backdrop(size: Tuple[int, int], background_path: str,
                     blur: int = 12, darkness: float = 0.34,
                     scrim_height: int = 150) -> Image.Image:
    """生成最底层装饰背景：花瓣 / 小电脑 / 软盘散落，再按旧背景同样的方式模糊压暗。"""
    width, height = max(1, int(size[0])), max(1, int(size[1]))
    key = (width, height)
    cached = _AMBIENT_CACHE.get(key)
    if cached is not None:
        _AMBIENT_CACHE.move_to_end(key)
        return cached

    deep, light, blue = _ambient_colors(background_path)
    colors = (light, deep, blue)
    layer = _ambient_base((width, height))
    base_w, base_h = AMBIENT_BASE_SIZE
    sx = width / float(base_w)
    sy = height / float(base_h)
    scale = min(sx, sy)
    for kind, cx, cy, size, angle, color_index in _ambient_layout():
        sprite = _shape_sprite(kind, max(12, int(round(size * scale))), colors[color_index], angle)
        layer.paste(sprite, (int(round(cx * sx)) - sprite.width // 2,
                             int(round(cy * sy)) - sprite.height // 2), sprite)

    image = layer.filter(ImageFilter.GaussianBlur(blur))
    image = ImageEnhance.Brightness(image).enhance(1.08)
    # 图形本身比旧照片背景更暗，压暗系数降低，保证边缘处仍能看出装饰形状。
    tint = Image.new("RGB", image.size, (12, 30, 20))
    image = Image.blend(image, tint, darkness)

    # 标题区压暗：只压文字所在的左右两块，页眉中部的空白留亮，
    # 让整窗仍能看见散落的装饰图形；两块压暗区边缘做高斯柔化，避免硬边。
    if scrim_height > 0 and height > scrim_height:
        band = image.crop((0, 0, width, scrim_height))
        mask = Image.new("L", (width, scrim_height), 0)
        draw = ImageDraw.Draw(mask)
        draw.rectangle((0, 0, min(790, width), 96), fill=255)
        if width > 1000:
            draw.rectangle((max(1000, width - 190), 0, width, 96), fill=255)
        mask = mask.filter(ImageFilter.GaussianBlur(18))
        factor = mask.point(lambda value: int(round(255 * (1.0 - 0.55 * value / 255.0))))
        band = ImageChops.multiply(band, Image.merge("RGB", (factor, factor, factor)))
        image.paste(band, (0, 0))

    _AMBIENT_CACHE[key] = image
    while len(_AMBIENT_CACHE) > _AMBIENT_CACHE_LIMIT:
        _AMBIENT_CACHE.popitem(last=False)
    return image


def load_character(image_path: str, padding: int = 2) -> Image.Image:
    """从背景图里裁出人物本体。

    background.png 自带 alpha 通道，这里只按 alpha 包围盒裁掉四周留白，
    保留原始透明边缘，不会带上矩形色块。
    """
    cached = _CHARACTER_MASTER_CACHE.get(image_path)
    if cached is not None:
        return cached
    with Image.open(image_path) as src:
        im = src.convert("RGBA")
    bbox = im.getchannel("A").getbbox()
    if bbox:
        left = max(0, bbox[0] - padding)
        top = max(0, bbox[1] - padding)
        right = min(im.width, bbox[2] + padding)
        bottom = min(im.height, bbox[3] + padding)
        im = im.crop((left, top, right, bottom))
    _CHARACTER_MASTER_CACHE[image_path] = im
    return im


def character_image(image_path: str, max_height: int, max_width: int,
                    opacity: float = 1.0) -> Image.Image:
    """按比例缩放人物立绘并整体调整透明度。

    尺寸与透明度组合带缓存：窗口反复重绘时不会每次都做 LANCZOS 重采样。
    """
    master = load_character(image_path)
    max_height = max(1, int(max_height))
    max_width = max(1, int(max_width))
    scale = min(max_height / float(master.height), max_width / float(master.width))
    width = max(1, int(round(master.width * scale)))
    height = max(1, int(round(master.height * scale)))
    alpha = int(round(max(0.0, min(1.0, opacity)) * 255))
    key = (image_path, width, height, alpha)
    cached = _CHARACTER_SCALED_CACHE.get(key)
    if cached is not None:
        _CHARACTER_SCALED_CACHE.move_to_end(key)
        return cached
    im = master.resize((width, height), Image.LANCZOS)
    if alpha < 255:
        im = im.copy()
        im.putalpha(im.getchannel("A").point(lambda value: value * alpha // 255))
    _CHARACTER_SCALED_CACHE[key] = im
    while len(_CHARACTER_SCALED_CACHE) > _CHARACTER_CACHE_LIMIT:
        _CHARACTER_SCALED_CACHE.popitem(last=False)
    return im
