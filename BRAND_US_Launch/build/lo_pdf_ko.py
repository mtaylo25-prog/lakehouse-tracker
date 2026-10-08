#!/usr/bin/env python3
"""PDF export for Korean decks through LibreOffice UNO.

LibreOffice applies "spacing between Asian and non Asian text" (ParaIsCharacterDistance) to every imported
paragraph, which puts visible gaps between Hangul and Latin text and before periods in the PDF. PowerPoint does
not do this. This script opens the pptx, switches the property off on every text shape and table cell (slides,
masters and layouts), then exports the PDF with the same filter to_pdf uses.

Usage: python3 build/lo_pdf_ko.py deck.pptx [out.pdf]
"""
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, "/mnt/skills/public/pptx/scripts")
from office.soffice import get_soffice_env  # noqa: E402

import uno  # noqa: E402
from com.sun.star.beans import PropertyValue  # noqa: E402
from com.sun.star.connection import NoConnectException  # noqa: E402

PORT = 2083


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def start_office(profile):
    args = ["soffice", f"-env:UserInstallation={Path(profile).as_uri()}", "--headless", "--invisible", "--norestore",
            "--nologo", f"--accept=socket,host=127.0.0.1,port={PORT};urp;"]
    return subprocess.Popen(args, env=get_soffice_env(), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def connect(timeout=90):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    t0 = time.time()
    while True:
        try:
            return resolver.resolve(f"uno:socket,host=127.0.0.1,port={PORT};urp;StarOffice.ComponentContext")
        except NoConnectException:
            if time.time() - t0 > timeout:
                raise
            time.sleep(0.5)


PARA_PROPS = (("ParaIsCharacterDistance", False),   # no gap between Hangul and Latin text or before punctuation
              ("ParaIsForbiddenRules", True))        # Korean line rules: no line ends with "(" or starts with ")" or "."


def fix_text(obj, stats):
    """Set the Korean paragraph properties on one text bearing object (shape or table cell)."""
    try:
        for name, value in PARA_PROPS:
            obj.setPropertyValue(name, value)
        stats["shapes"] += 1
    except Exception:
        pass
    try:
        enum = obj.getText().createEnumeration() if hasattr(obj, "getText") else obj.createEnumeration()
        while enum.hasMoreElements():
            para = enum.nextElement()
            try:
                for name, value in PARA_PROPS:
                    para.setPropertyValue(name, value)
                stats["paras"] += 1
            except Exception:
                pass
    except Exception:
        pass


def walk_shape(shape, stats):
    st = getattr(shape, "ShapeType", "")
    if st == "com.sun.star.drawing.GroupShape":
        for i in range(shape.getCount()):
            walk_shape(shape.getByIndex(i), stats)
        return
    if st == "com.sun.star.drawing.TableShape":
        model = shape.Model
        for r in range(model.RowCount):
            for c in range(model.ColumnCount):
                fix_text(model.getCellByPosition(c, r), stats)
        return
    fix_text(shape, stats)


def walk_pages(pages, stats):
    for i in range(pages.getCount()):
        page = pages.getByIndex(i)
        for j in range(page.getCount()):
            walk_shape(page.getByIndex(j), stats)


def export(src, dst):
    src, dst = Path(src).resolve(), Path(dst).resolve()
    with tempfile.TemporaryDirectory(prefix="lo_profile_ko_") as profile:
        office = start_office(profile)
        try:
            ctx = connect()
            desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            doc = desktop.loadComponentFromURL(src.as_uri(), "_blank", 0, (prop("Hidden", True),))
            stats = {"shapes": 0, "paras": 0}
            walk_pages(doc.getDrawPages(), stats)
            walk_pages(doc.getMasterPages(), stats)
            doc.storeToURL(dst.as_uri(), (prop("FilterName", "impress_pdf_Export"),))
            doc.close(True)
            print(f"wrote {dst} (autospacing off on {stats['shapes']} shapes or cells, {stats['paras']} paragraphs)")
            try:
                desktop.terminate()
            except Exception:
                pass
        finally:
            try:
                office.wait(timeout=20)
            except Exception:
                office.kill()


if __name__ == "__main__":
    s = sys.argv[1]
    d = sys.argv[2] if len(sys.argv) > 2 else str(Path(s).with_suffix(".pdf"))
    export(s, d)
