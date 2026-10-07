"""
Crée un raccourci "BookBinder" dans le menu Démarrer de l'utilisateur.
Lance : python install_shortcut.py
"""
import os
import sys


def _pause():
    try:
        if sys.stdin and sys.stdin.isatty():
            input("\nAppuie sur Entrée pour fermer...")
    except EOFError:
        pass


def render_icon_image(S=1024):
    """Dessine l'icône en haute résolution : double page de BD sur fond bleu clair."""
    from PIL import Image, ImageChops, ImageDraw, ImageFilter

    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

    # Fond : carré arrondi avec dégradé vertical bleu clair
    grad = Image.new("RGBA", (S, S))
    gd = ImageDraw.Draw(grad)
    top, bot = (150, 208, 248), (52, 128, 202)
    for y in range(S):
        t = y / (S - 1)
        color = tuple(int(top[i] + (bot[i] - top[i]) * t) for i in range(3))
        gd.line([(0, y), (S, y)], fill=color + (255,))
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle([28, 28, S - 28, S - 28],
                                           radius=230, fill=255)
    img.paste(grad, (0, 0), mask)

    # Reflet doux en haut du fond
    gloss = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(gloss).ellipse([-220, -460, S + 220, int(S * 0.40)],
                                  fill=(255, 255, 255, 42))
    gloss.putalpha(ImageChops.multiply(gloss.getchannel("A"), mask))
    img = Image.alpha_composite(img, gloss)

    # Ombre portée sous les pages
    shadow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle([206, 280, 818, 812],
                                             radius=36, fill=(10, 44, 88, 140))
    shadow = shadow.filter(ImageFilter.GaussianBlur(26))
    img = Image.alpha_composite(img, shadow)

    d = ImageDraw.Draw(img)

    # Deux pages blanches côte à côte (double page), séparées par la reliure
    cx, gap = S // 2, 9
    d.rounded_rectangle([206, 252, cx - gap, 788], radius=34, fill=(255, 255, 255, 255))
    d.rounded_rectangle([cx + gap, 252, 818, 788], radius=34, fill=(255, 255, 255, 255))

    # Cases de BD sur chaque page, en bleu clair
    panel, panel2 = (158, 206, 244, 255), (192, 226, 250, 255)
    r = 14
    # Page gauche : grande case en haut, deux petites en bas
    d.rounded_rectangle([242, 292, 468, 486], radius=r, fill=panel)
    d.rounded_rectangle([242, 518, 342, 748], radius=r, fill=panel2)
    d.rounded_rectangle([368, 518, 468, 748], radius=r, fill=panel2)
    # Page droite : deux petites en haut, grande case en bas (miroir)
    d.rounded_rectangle([556, 292, 656, 522], radius=r, fill=panel2)
    d.rounded_rectangle([682, 292, 782, 522], radius=r, fill=panel2)
    d.rounded_rectangle([556, 554, 782, 748], radius=r, fill=panel)

    return img


def create_icon_file(icon_path):
    from PIL import Image
    base = render_icon_image(1024)
    sizes = [16, 24, 32, 48, 64, 128, 256]
    frames = [base.resize((s, s), Image.LANCZOS) for s in sizes]
    # La première frame doit être la plus grande : Pillow ignore les tailles
    # supérieures à l'image de base (c'était la cause de l'icône pixelisée).
    frames[-1].save(icon_path, format="ICO",
                    sizes=[(s, s) for s in sizes],
                    append_images=frames[:-1])


def main():
    app_dir = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(app_dir, "fusion_images.py")

    if not os.path.exists(script_path):
        print(f"❌ fusion_images.py introuvable dans : {app_dir}")
        _pause()
        return

    # pythonw.exe = Python sans fenêtre console
    python_dir = os.path.dirname(sys.executable)
    pythonw = os.path.join(python_dir, "pythonw.exe")
    if not os.path.exists(pythonw):
        pythonw = sys.executable

    # Icône .ico
    icon_path = os.path.join(app_dir, "fusion_images.ico")
    print("🎨 Génération de l'icône...")
    try:
        create_icon_file(icon_path)
        print(f"   ✅ Icône créée : {icon_path}")
    except Exception as e:
        print(f"   ⚠️  Icône ignorée ({e}), le raccourci utilisera l'icône Python par défaut")
        icon_path = pythonw

    # Dossier menu Démarrer utilisateur
    start_menu = os.path.join(
        os.environ.get("APPDATA", ""),
        "Microsoft", "Windows", "Start Menu", "Programs"
    )
    shortcut_path = os.path.join(start_menu, "BookBinder.lnk")

    print("📌 Création du raccourci menu Démarrer...")
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(shortcut_path)
        shortcut.TargetPath = pythonw
        shortcut.Arguments = f'"{script_path}"'
        shortcut.WorkingDirectory = app_dir
        shortcut.IconLocation = icon_path
        shortcut.Description = "BookBinder — Fusion de doubles pages BD"
        shortcut.WindowStyle = 1
        shortcut.Save()
        print(f"   ✅ Raccourci créé : {shortcut_path}")
        print("\n🌸 BookBinder est maintenant dans ton menu Démarrer !")
    except ImportError:
        print("❌ pywin32 manquant — pip install pywin32")
    except Exception as e:
        print(f"❌ Erreur : {e}")

    _pause()


if __name__ == "__main__":
    main()
