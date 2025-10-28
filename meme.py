from io import BytesIO
from PIL import Image, ImageFont, ImageDraw
import argparse
import emoji
import os
import struct
import sys


class Win32:
    if sys.platform != 'win32':
        raise SystemError('This module is only for Windows.')

    import ctypes
    import ctypes.wintypes as w

    user32 = ctypes.WinDLL('user32')
    kernel32 = ctypes.WinDLL('kernel32')

    OpenClipboard = user32.OpenClipboard
    OpenClipboard.argtypes = w.HWND,
    OpenClipboard.restype = w.BOOL

    EmptyClipboard = user32.EmptyClipboard
    EmptyClipboard.argtypes = None
    EmptyClipboard.restype = w.BOOL

    _SetClipboardData = user32.SetClipboardData
    _SetClipboardData.argtypes = w.UINT,w.HANDLE,
    _SetClipboardData.restype = w.HANDLE

    GlobalAlloc = kernel32.GlobalAlloc
    GlobalAlloc.argtypes = w.UINT,ctypes.c_size_t,
    GlobalAlloc.restype = w.HGLOBAL

    GlobalLock = kernel32.GlobalLock
    GlobalLock.argtypes = w.HGLOBAL,
    GlobalLock.restype = w.LPVOID

    GlobalUnlock = kernel32.GlobalUnlock
    GlobalUnlock.argtypes = w.HGLOBAL,
    GlobalUnlock.restype = w.BOOL

    CloseClipboard = user32.CloseClipboard
    CloseClipboard.argtypes = None

    RegisterClipboardFormat = user32.RegisterClipboardFormatW
    RegisterClipboardFormat.argtypes = w.LPCWSTR,
    RegisterClipboardFormat.restype = w.UINT

    CF_DIB = 8

    def SetClipboardData(format, data):
        if isinstance(data, int):
            return Win32.SetClipboardData(format, struct.pack('@I', data))
        elif isinstance(data, bytes):
            GMEM_MOVEABLE = 0x0002
            GMEM_ZEROINIT = 0x0040
            size = len(data)
            clipboardData = Win32.GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, size)
            pchData = Win32.GlobalLock(clipboardData)

            try:
                buffer = Win32.ctypes.create_string_buffer(data, size)
                Win32.ctypes.memmove(pchData, buffer, size)
                return Win32._SetClipboardData(format, pchData)
            finally:
                Win32.GlobalUnlock(clipboardData)
        else:
            raise TypeError(f'Unsupported data type: {type(data)}')


class StringIgnoreCase(str):
    def __eq__(self, other: str) -> bool:
        return self.casefold() == other.casefold()


def get_clipboard_format(image_type: StringIgnoreCase):
    if image_type == 'png':
        return Win32.RegisterClipboardFormat('PNG')
    elif image_type == 'bmp':
        return Win32.CF_DIB


def send_to_clipboard(image_type: StringIgnoreCase, data):
    Win32.OpenClipboard(None)
    try:
        Win32.EmptyClipboard()

        # exclude from clipboard history and cloud clipboard
        hist = Win32.RegisterClipboardFormat('CanIncludeInClipboardHistory')
        cloud = Win32.RegisterClipboardFormat('CanUploadToCloudClipboard')
        Win32.SetClipboardData(hist, 0)
        Win32.SetClipboardData(cloud, 0)
        Win32.SetClipboardData(get_clipboard_format(image_type), data)
    finally:
        Win32.CloseClipboard()


# Segment represents either a text or an emoji segment
def get_segments(text):
    segments = []
    offset = 0
    for t in emoji.analyze(text):
        if offset != t.value.start:
            segments.append((text[offset:t.value.start], False))

        segments.append((t.value.emoji, True))
        offset = t.value.end

    if last_segment := text[offset:]:
        segments.append((last_segment, False))

    return segments


def draw_meme_text(image, line1, line2):
    image_draw = ImageDraw.Draw(image)

    font_size = max(20, image.height // 10)
    font_text = ImageFont.truetype('impact.ttf', font_size, encoding='unic')
    font_emoji = ImageFont.truetype('seguiemj.ttf', int(font_size), encoding='unic')

    textbbox_args = {
        'stroke_width': max(2, font_size // 15),
    }
    draw_args = textbbox_args | {
        'stroke_fill': (0, 0, 0),
    }

    text_color = (255, 255, 255)
    margin = image.height * 0.04

    if line2 is not None:
        top = line1
        bottom = line2
    elif line1:
        top = None
        bottom = line1
    else:
        top = bottom = None

    TOP = 1
    BOTTOM = 2

    def draw_emoji_text(position, y, text):
        parts = []
        for segment, is_emoji in get_segments(text):
            if is_emoji:
                font = font_emoji
                args = draw_args | {
                    'font': font_emoji,
                    'embedded_color': True,
                }
            else:
                font = font_text
                args = draw_args | {
                    'font': font_text,
                }

            _, _, w, h = image_draw.textbbox((0, 0), segment, font=font, **textbbox_args)
            parts.append((segment, args, w, h))

        total_width = sum(w for _, _, w, _ in parts)
        max_h = max(h for _, _, _, h in parts)
        x = (image.width - total_width) / 2

        for text, args, w, h in parts:
            if position == BOTTOM:
                text_y = y - max_h
            else:
                text_y = y

            text_y += max_h - h
            image_draw.text((x, text_y), text, text_color, **args)
            x += w

    if bottom:
        draw_emoji_text(BOTTOM, image.height - margin, bottom)

    if top:
        draw_emoji_text(TOP, margin, top)


def save_meme_to_clipboard(image_path, line1, line2, image_type: StringIgnoreCase):
    image = Image.open(image_path)
    draw_meme_text(image, line1=line1, line2=line2)

    if image_type == 'bmp':
        with BytesIO() as output:
            image.convert("RGB").save(output, 'BMP')
            data = output.getvalue()[14:]
    elif image_type == 'png':
        with BytesIO() as output:
            image.convert("RGB").save(output, 'PNG')
            data = output.getvalue()
    else:
        assert False, f'Unsupported image type: {image_type}'

    send_to_clipboard(image_type, data)


def main():
    default_image = os.path.join(os.environ['OneDriveConsumer'], 'Pictures', 'Merci.png')
    parser = argparse.ArgumentParser(description='Meme Generator')
    parser.add_argument('line1', type=str, help='line 1')
    parser.add_argument('line2', nargs='?', type=str, help='line 2')
    parser.add_argument('-i', '--image', type=str, default=default_image, help='Image file path (default: %(default)s)')

    output_group = parser.add_mutually_exclusive_group()
    output_group.add_argument('-t', '--type', type=StringIgnoreCase, default='bmp', choices=('bmp', 'png'), help='Type of image to save in the clipboard (default: %(default)s)')
    output_group.add_argument('-o', '--output', type=str, help='Save image to specified file instead of the clipboard')

    args = parser.parse_args()

    if args.output:
        image = Image.open(args.image)
        draw_meme_text(image, line1=args.line1, line2=args.line2)
        image.save(args.output)
    else:
        save_meme_to_clipboard(args.image, args.line1, args.line2, args.type)


if __name__ == '__main__':
    main()
