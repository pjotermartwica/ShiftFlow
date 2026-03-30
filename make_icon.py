"""Generuje icon.ico dla aplikacji ShiftFlow."""
from PIL import Image, ImageDraw, ImageFont
import os

def make_icon():
    sizes = [256, 128, 64, 48, 32, 16]
    images = []

    for size in sizes:
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Tło – zaokrąglony prostokąt (ciemny granat)
        margin = max(1, size // 16)
        r = size // 6
        bg = (28, 40, 65, 255)
        accent = (0, 200, 130, 255)   # szmaragdowy
        header = (0, 150, 100, 255)

        # Prostokąt tła
        draw.rounded_rectangle([margin, margin, size - margin, size - margin],
                                radius=r, fill=bg)

        # Pasek nagłówka kalendarza
        hh = max(4, size // 6)
        draw.rounded_rectangle([margin, margin, size - margin, margin + hh],
                                radius=r, fill=header)

        # Uchwyty kalendarza (dwa prostokąty u góry)
        hook_w = max(2, size // 14)
        hook_h = max(3, size // 9)
        hook_y = margin - hook_h // 2
        lx = size // 4
        rx = size * 3 // 4
        for hx in [lx, rx]:
            draw.rounded_rectangle(
                [hx - hook_w // 2, hook_y, hx + hook_w // 2, hook_y + hook_h],
                radius=max(1, hook_w // 2), fill=accent
            )

        # Siatka komórek (3 kolumny × 3 wiersze)
        if size >= 32:
            cols, rows = 3, 3
            pad = max(2, size // 12)
            cell_area_x1 = margin + pad
            cell_area_y1 = margin + hh + pad
            cell_area_x2 = size - margin - pad
            cell_area_y2 = size - margin - pad
            cw = (cell_area_x2 - cell_area_x1) // cols
            ch = (cell_area_y2 - cell_area_y1) // rows
            for row in range(rows):
                for col in range(cols):
                    cx1 = cell_area_x1 + col * cw + 1
                    cy1 = cell_area_y1 + row * ch + 1
                    cx2 = cx1 + cw - 3
                    cy2 = cy1 + ch - 3
                    if cx2 > cx1 and cy2 > cy1:
                        color = accent if (row + col) % 3 == 0 else (60, 80, 110, 200)
                        draw.rounded_rectangle([cx1, cy1, cx2, cy2],
                                               radius=max(1, size // 40), fill=color)

        images.append(img)

    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
    images[0].save(out_path, format="ICO", sizes=[(s, s) for s in sizes],
                   append_images=images[1:])
    print(f"Ikona zapisana: {out_path}")

if __name__ == "__main__":
    make_icon()
