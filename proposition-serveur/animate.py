"""Ajoute transitions et animations d'entrée (fondu + montée) à la présentation.

Les formes nommées « aN_... » apparaissent automatiquement, groupe par groupe (N croissant).
Usage : python3 animate.py deck.pptx
"""
import re
import sys
import zipfile

STEP_MS, STAGGER_MS, DUR_MS = 450, 60, 600


def effect(cid, spid, delay):
    tgt = f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
    return (
        f'<p:par><p:cTn id="{cid}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect">'
        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
        f'<p:set><p:cBhvr><p:cTn id="{cid + 1}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}'
        f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
        f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{cid + 2}" dur="{DUR_MS}"/>{tgt}</p:cBhvr></p:animEffect>'
        f'<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="{cid + 3}" dur="{DUR_MS}" decel="100000" fill="hold"/>{tgt}'
        f'<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>'
        f'<p:tav tm="0"><p:val><p:strVal val="#ppt_y+0.03"/></p:val></p:tav>'
        f'<p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>'
        f'</p:childTnLst></p:cTn></p:par>'
    )


def timing(shapes):
    cid, parts, last = 4, [], None
    idx = 0
    for spid, grp in shapes:
        idx = idx + 1 if grp == last else 0
        last = grp
        parts.append(effect(cid, spid, (grp - 1) * STEP_MS + idx * STAGGER_MS))
        cid += 4
    return (
        '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
        '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
        + ''.join(parts) +
        '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
    )


def process(xml):
    shapes = [(int(i), int(g)) for i, g in re.findall(r'<p:cNvPr id="(\d+)" name="a(\d+)_', xml)]
    shapes.sort(key=lambda s: s[1])
    add = '<p:transition spd="slow"><p:fade/></p:transition>'
    if shapes:
        add += timing(shapes)
    xml = re.sub(r'<p:transition.*?</p:transition>|<p:timing>.*?</p:timing>', '', xml, flags=re.S)
    if '</p:clrMapOvr>' in xml:
        return xml.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + add, 1)
    return xml.replace('</p:cSld>', '</p:cSld>' + add, 1) if '<p:clrMapOvr' not in xml else xml


def main(path):
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if re.fullmatch(r'ppt/slides/slide\d+\.xml', info.filename):
                data = process(data.decode('utf-8')).encode('utf-8')
            z.writestr(info, data)


if __name__ == '__main__':
    main(sys.argv[1])
