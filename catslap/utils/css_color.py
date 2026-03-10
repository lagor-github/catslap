import math
from catslap.utils import text as text_util

__COLOR_MAPPING = {
    'magenta': '#ff00ff',
    'fuchsia': '#ff00ff',
    'gray': '#808080',
    'darkred': '#8b0000',
    'brown': '#a52a2a',
    'firebrick': '#b22222',
    'crimson': '#dc143c',
    'red': '#ff0000',
    'tomato': '#ff6347',
    'coral': '#ff7f50',
    'indianred': '#cd5c5c',
    'lightcoral': '#f08080',
    'darksalmon': '#e9967a',
    'salmon': '#fa8072',
    'lightsalmon': '#ffa07a',
    'orangered': '#ff4500',
    'darkorange': '#ff8c00',
    'orange': '#ffa500',
    'gold': '#ffd700',
    'darkgoldenrod': '#b8860b',
    'goldenrod': '#daa520',
    'palegoldenrod': '#eee8aa',
    'darkkhaki': '#bdb76b',
    'khaki': '#f0e68c',
    'olive': '#808000',
    'yellow': '#ffff00',
    'yellowgreen': '#9acd32',
    'darkolivegreen': '#556b2f',
    'olivedrab': '#6b8e23',
    'lawngreen': '#7cfc00',
    'chartreuse': '#7fff00',
    'greenyellow': '#adff2f',
    'darkgreen': '#006400',
    'green': '#008000',
    'forestgreen': '#228b22',
    'lime': '#00ff00',
    'limegreen': '#32cd32',
    'lightgreen': '#90ee90',
    'palegreen': '#98fb98',
    'darkseagreen': '#8fbc8f',
    'mediumspringgreen': '#00fa9a',
    'springgreen': '#00ff7f',
    'seagreen': '#2e8b57',
    'mediumaquamarine': '#66cdaa',
    'mediumseagreen': '#3cb371',
    'lightseagreen': '#20b2aa',
    'darkslategray': '#2f4f4f',
    'teal': '#008080',
    'darkcyan': '#008b8b',
    'aqua': '#00ffff',
    'cyan': '#00ffff',
    'lightcyan': '#e0ffff',
    'darkturquoise': '#00ced1',
    'turquoise': '#40e0d0',
    'mediumturquoise': '#48d1cc',
    'paleturquoise': '#afeeee',
    'aquamarine': '#7fffd4',
    'powderblue': '#b0e0e6',
    'cadetblue': '#5f9ea0',
    'steelblue': '#4682b4',
    'cornflowerblue': '#6495ed',
    'deepskyblue': '#00bfff',
    'dodgerblue': '#1e90ff',
    'lightblue': '#add8e6',
    'skyblue': '#87ceeb',
    'lightskyblue': '#87cefa',
    'midnightblue': '#191970',
    'navy': '#000080',
    'darkblue': '#00008b',
    'mediumblue': '#0000cd',
    'blue': '#0000ff',
    'royalblue': '#4169e1',
    'blueviolet': '#8a2be2',
    'indigo': '#4b0082',
    'darkslateblue': '#483d8b',
    'slateblue': '#6a5acd',
    'mediumslateblue': '#7b68ee',
    'mediumpurple': '#9370db',
    'darkmagenta': '#8b008b',
    'darkviolet': '#9400d3',
    'darkorchid': '#9932cc',
    'mediumorchid': '#ba55d3',
    'purple': '#800080',
    'thistle': '#d8bfd8',
    'plum': '#dda0dd',
    'violet': '#ee82ee',
    'magenta/fuchsia': '#ff00ff',
    'orchid': '#da70d6',
    'mediumvioletred': '#c71585',
    'palevioletred': '#db7093',
    'deeppink': '#ff1493',
    'hotpink': '#ff69b4',
    'lightpink': '#ffb6c1',
    'pink': '#ffc0cb',
    'antiquewhite': '#faebd7',
    'beige': '#f5f5dc',
    'bisque': '#ffe4c4',
    'blanchedalmond': '#ffebcd',
    'wheat': '#f5deb3',
    'cornsilk': '#fff8dc',
    'lemonchiffon': '#fffacd',
    'lightgoldenrodyellow': '#fafad2',
    'lightyellow': '#ffffe0',
    'saddlebrown': '#8b4513',
    'sienna': '#a0522d',
    'chocolate': '#d2691e',
    'peru': '#cd853f',
    'sandybrown': '#f4a460',
    'burlywood': '#deb887',
    'tan': '#d2b48c',
    'rosybrown': '#bc8f8f',
    'moccasin': '#ffe4b5',
    'navajowhite': '#ffdead',
    'peachpuff': '#ffdab9',
    'mistyrose': '#ffe4e1',
    'lavenderblush': '#fff0f5',
    'linen': '#faf0e6',
    'oldlace': '#fdf5e6',
    'papayawhip': '#ffefd5',
    'seashell': '#fff5ee',
    'mintcream': '#f5fffa',
    'slategray': '#708090',
    'lightslategray': '#778899',
    'lightsteelblue': '#b0c4de',
    'lavender': '#e6e6fa',
    'floralwhite': '#fffaf0',
    'aliceblue': '#f0f8ff',
    'ghostwhite': '#f8f8ff',
    'honeydew': '#f0fff0',
    'ivory': '#fffff0',
    'azure': '#f0ffff',
    'snow': '#fffafa',
    'black': '#000000',
    'dimgray': '#696969',
    'dimgrey': '#696969',
    'grey': '#808080',
    'darkgray': '#a9a9a9',
    'darkgrey': '#a9a9a9',
    'silver': '#c0c0c0',
    'lightgray': '#d3d3d3',
    'lightgrey': '#d3d3d3',
    'gainsboro': '#dcdcdc',
    'whitesmoke': '#f5f5f5',
    'white': '#ffffff',
}

def get_rgb_color(color: str|None) -> str|None:
  if not color:
    return None
  color = color.strip().lower()
  mapped_color = __COLOR_MAPPING.get(color)
  if mapped_color:
    color = mapped_color[1:]
  else:
    try:
      r, g, b, a = _parse_css_color(color)
      r, g, b = _blend_white(r, g, b, a)
      color = f"{int(r):02X}{int(g):02X}{int(b):02X}"
    except Exception:
      color = "000000"
  return color

def _blend_white(r, g, b, a):
  r = r * a + 255 * (1 - a)
  g = g * a + 255 * (1 - a)
  b = b * a + 255 * (1 - a)
  return round(r), round(g), round(b)

def _parse_css_color(c):
  if c.startswith("#"):
    return _parse_hex(c)
  if c.startswith(("rgb(", "rgba(")):
    return _parse_rgb(c)
  if c.startswith(("hsl(", "hsla(")):
    return _parse_hsl(c)
  if c.startswith("hwb("):
    return _parse_hwb(c)
  if c.startswith("lab("):
    return _parse_lab(c)
  if c.startswith("lch("):
    return _parse_lch(c)
  if c.startswith("oklab("):
    return _parse_oklab(c)
  if c.startswith("oklch("):
    return _parse_oklch(c)
  if c.startswith("color("):
    return _parse_color_function(c)
  if text_util.is_hex(c):
    return _parse_hex('#' + c);
  raise ValueError("Cannot parse value: " + str(c))

# --------------------------------------------------
# HEX
# --------------------------------------------------
def _parse_hex(c):
  h = c[1:]
  if len(h) == 3:
    r, g, b = [int(x * 2, 16) for x in h]
    return r, g, b, 1
  if len(h) == 4:
    r, g, b, a = [int(x * 2, 16) for x in h]
    return r, g, b, a / 255
  if len(h) == 6:
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return r, g, b, 1
  if len(h) == 8:
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    a = int(h[6:8], 16) / 255
    return r, g, b, a
  raise ValueError


# --------------------------------------------------
# RGB
# --------------------------------------------------
def _parse_rgb(c):
  c = c[c.find("(") + 1:c.rfind(")")]
  c = c.replace(",", " ")
  if "/" in c:
    main, alpha = c.split("/")
    a = _parse_alpha(alpha.strip())
  else:
    main = c
    a = 1
  parts = [p for p in main.split() if p]
  r = _parse_rgb_value(parts[0])
  g = _parse_rgb_value(parts[1])
  b = _parse_rgb_value(parts[2])
  return r, g, b, a

def _parse_rgb_value(v):
  if v.endswith("%"):
    return float(v[:-1]) * 255 / 100
  return float(v)

def _parse_alpha(v):
  v = v.strip()
  if v.endswith("%"):
    return float(v[:-1]) / 100
  return float(v)

# --------------------------------------------------
# HSL
# --------------------------------------------------
def _parse_hsl(c):
  c = c[c.find("(") + 1:c.rfind(")")]
  c = c.replace(",", " ")
  if "/" in c:
    main, alpha = c.split("/")
    a = _parse_alpha(alpha.strip())
  else:
    main = c
    a = 1
  h, s, l = [p for p in main.split() if p][:3]
  h = _parse_angle(h)
  s = float(s.rstrip("%")) / 100
  l = float(l.rstrip("%")) / 100
  r, g, b = _hsl_to_rgb(h, s, l)
  return r * 255, g * 255, b * 255, a

def _hsl_to_rgb(h, s, l):
  c = (1 - abs(2 * l - 1)) * s
  x = c * (1 - abs((h / 60) % 2 - 1))
  m = l - c / 2
  if h < 60:
    r, g, b = c, x, 0
  elif h < 120:
    r, g, b = x, c, 0
  elif h < 180:
    r, g, b = 0, c, x
  elif h < 240:
    r, g, b = 0, x, c
  elif h < 300:
    r, g, b = x, 0, c
  else:
    r, g, b = c, 0, x
  return r + m, g + m, b + m

# --------------------------------------------------
# HWB
# --------------------------------------------------
def _parse_hwb(c):
  c = c[c.find("(") + 1:c.rfind(")")]
  c = c.replace(",", " ")
  if "/" in c:
    main, alpha = c.split("/")
    a = _parse_alpha(alpha)
  else:
    main = c
    a = 1
  h, w, b = main.split()[:3]
  h = _parse_angle(h)
  w = float(w.rstrip("%")) / 100
  b = float(b.rstrip("%")) / 100
  r, g, bl = _hsl_to_rgb(h, 1, 0.5)
  r = r * (1 - w - b) + w
  g = g * (1 - w - b) + w
  bl = bl * (1 - w - b) + w
  return r * 255, g * 255, bl * 255, a


# --------------------------------------------------
# LAB / LCH
# --------------------------------------------------
def _parse_lab(c):
  c = c[c.find("(") + 1:c.rfind(")")]
  parts = c.replace("/", " ").split()
  L = float(parts[0].rstrip("%"))
  a = float(parts[1])
  b = float(parts[2])
  alpha = 1
  if len(parts) == 4:
    alpha = _parse_alpha(parts[3])
  return _lab_to_rgb(L, a, b, alpha)


def _parse_lch(c):
  c = c[c.find("(") + 1:c.rfind(")")]
  parts = c.replace("/", " ").split()
  L = float(parts[0].rstrip("%"))
  C = float(parts[1])
  h = math.radians(_parse_angle(parts[2]))
  a = C * math.cos(h)
  b = C * math.sin(h)
  alpha = 1
  if len(parts) == 4:
    alpha = _parse_alpha(parts[3])
  return _lab_to_rgb(L, a, b, alpha)


# --------------------------------------------------
# OKLAB / OKLCH (simplificado)
# --------------------------------------------------
def _parse_oklab(c):
  c = c[c.find("(") + 1:c.rfind(")")]
  parts = c.replace("/", " ").split()
  L, a, b = map(float, parts[:3])
  alpha = 1
  if len(parts) == 4:
      alpha = _parse_alpha(parts[3])
  return _oklab_to_rgb(L, a, b, alpha)


def _parse_oklch(c):
  c = c[c.find("(") + 1:c.rfind(")")]
  parts = c.replace("/", " ").split()
  L = float(parts[0])
  C = float(parts[1])
  h = math.radians(_parse_angle(parts[2]))
  a = C * math.cos(h)
  b = C * math.sin(h)
  alpha = 1
  if len(parts) == 4:
      alpha = _parse_alpha(parts[3])
  return _oklab_to_rgb(L, a, b, alpha)


# --------------------------------------------------
# color(srgb ...)
# --------------------------------------------------
def _parse_color_function(c):
  inside = c[c.find("(") + 1:c.rfind(")")]
  parts = inside.split()
  space = parts[0]
  if space != "srgb":
    raise ValueError
  r = float(parts[1]) * 255
  g = float(parts[2]) * 255
  b = float(parts[3]) * 255
  a = 1
  if "/" in inside:
    a = _parse_alpha(inside.split("/")[-1])
  return r, g, b, a

# --------------------------------------------------
# utilidades
# --------------------------------------------------
def _parse_angle(v):
  if v.endswith("deg"):
    return float(v[:-3])
  if v.endswith("turn"):
    return float(v[:-4]) * 360
  if v.endswith("rad"):
    return float(v[:-3]) * 180 / math.pi
  if v.endswith("grad"):
    return float(v[:-4]) * 0.9
  return float(v)

# --------------------------------------------------
# conversiones LAB y OKLAB → RGB (aprox)
# --------------------------------------------------
def _lab_to_rgb(L, a, b, alpha):
  y = (L + 16) / 116
  x = a / 500 + y
  z = y - b / 200
  def f(t):
    return t ** 3 if t ** 3 > 0.008856 else (t - 16 / 116) / 7.787

  x, y, z = f(x), f(y), f(z)
  x *= 95.047
  y *= 100
  z *= 108.883
  x /= 100
  y /= 100
  z /= 100
  r = x * 3.2406 + y * -1.5372 + z * -0.4986
  g = x * -0.9689 + y * 1.8758 + z * 0.0415
  b = x * 0.0557 + y * -0.2040 + z * 1.0570
  r = max(0, min(1, r))
  g = max(0, min(1, g))
  b = max(0, min(1, b))
  return r * 255, g * 255, b * 255, alpha


def _oklab_to_rgb(L, a, b, alpha):
  l = L + 0.3963377774 * a + 0.2158037573 * b
  m = L - 0.1055613458 * a - 0.0638541728 * b
  s = L - 0.0894841775 * a - 1.2914855480 * b
  l, m, s = l ** 3, m ** 3, s ** 3
  r = +4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s
  g = -1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s
  b = -0.0041960863 * l - 0.7034186147 * m + 1.7076147010 * s
  r = max(0, min(1, r))
  g = max(0, min(1, g))
  b = max(0, min(1, b))
  return r * 255, g * 255, b * 255, alpha
