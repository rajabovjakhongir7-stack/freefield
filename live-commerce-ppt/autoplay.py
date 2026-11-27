"""Make a video on one slide start automatically (PowerPoint "Start: Automatically").

Usage: python3 autoplay.py deck.pptx slide2.xml "<video objectName>" <duration_ms>
Adds a <p:timing> block that plays the named video when the slide opens.
"""
import re
import shutil
import sys
import tempfile
import zipfile

deck, slide, name, dur = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
part = f"ppt/slides/{slide}"

with zipfile.ZipFile(deck) as z:
    items = [(i, z.read(i.filename)) for i in z.infolist()]

out = []
for info, data in items:
    if info.filename == part:
        xml = data.decode("utf-8")
        m = re.search(r'<p:cNvPr id="(\d+)" name="' + re.escape(name) + '"', xml)
        if not m:
            sys.exit(f"autoplay: no shape named {name!r} in {part}")
        if "<p:timing>" in xml:
            sys.exit(f"autoplay: {part} already has <p:timing>")
        spid = m.group(1)
        timing = (
            '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:par><p:cTn id="2" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
            '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            '<p:par><p:cTn id="4" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">'
            '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="5" dur="{dur}" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:cmd>'
            '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
            '<p:video><p:cMediaNode vol="80000"><p:cTn id="6" fill="hold" display="0">'
            '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cMediaNode></p:video>'
            '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
        )
        # <p:timing> goes after cSld / clrMapOvr / transition and before extLst.
        if "<p:extLst>" in xml.split("</p:cSld>", 1)[1]:
            head, tail = xml.rsplit("<p:extLst>", 1)
            xml = head + timing + "<p:extLst>" + tail
        else:
            xml = xml.replace("</p:sld>", timing + "</p:sld>")
        data = xml.encode("utf-8")
    out.append((info, data))

tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx").name
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
    for info, data in out:
        z.writestr(info, data)
shutil.move(tmp, deck)
print(f"autoplay: {part} video spid={spid} set to start automatically")
