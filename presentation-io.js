// Reines Format-Modul ohne DOM: schreibt und liest PPTX- und ODP-Pakete.
//
// Geometrie:
//   - Alle Positionen sind Anteile 0..1 der Folienbreite/-höhe (x, y = linke obere Ecke).
//   - Schriftgrößen in Punkt (pt). Eine Folie entspricht 13,333 × 7,5 Zoll (960 × 540 pt).
//
// Export-Folie (Eingabe der Schreibfunktionen):
//   {
//     bg: "#eeece5" | null,                         // Hintergrundfarbe
//     layer: { mime: "image/jpeg", bytes: Uint8Array } | null, // Folie ohne Text als Bild
//     texts: [TextBox]                              // bearbeitbare Textfelder
//   }
//   TextBox: { text, x, y, w, h, fontPt, color, alpha, bold, italic,
//              family, align: "l"|"ctr"|"r"|"just", lineSpacingPt, spacingPt, wrap }
//
// Import-Folie (Ausgabe der Lesefunktionen):
//   { bg, elements: [ {type:"text", ...TextBox-ähnlich} | {type:"shape", shape, x, y, w, h, color}
//                     | {type:"image", x, y, w, h, mime, bytes} ] }

import JSZip from "jszip";

export const SLIDE_WIDTH_PT = 960;
export const SLIDE_HEIGHT_PT = 540;
const PPTX_W = 12192000; // EMU
const PPTX_H = 6858000;
const ODP_W_CM = 33.867;
const ODP_H_CM = 19.05;

const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  pkgRel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
};

const ODF = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  draw: "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
  xlink: "http://www.w3.org/1999/xlink",
  presentation: "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
  manifest: "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
};

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

// ---------- Hilfsfunktionen ----------

function esc(value) {
  return String(value)
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function hexOf(color, fallback = "#171717") {
  return /^#[0-9a-f]{6}$/i.test(color || "") ? color.toUpperCase() : fallback.toUpperCase();
}

function round(value, digits = 0) {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function alphaPercent(alpha) {
  return Math.round(clamp(alpha ?? 1, 0, 1) * 100000);
}

function pptxColor(color, alpha) {
  const a = alpha !== undefined && alpha < 1 ? `<a:alpha val="${alphaPercent(alpha)}"/>` : "";
  return `<a:solidFill><a:srgbClr val="${hexOf(color).slice(1)}">${a}</a:srgbClr></a:solidFill>`;
}

// Editor-Schriftstapel (z. B. "Georgia, serif") auf einen einzelnen Schriftnamen abbilden.
function plainFamily(family) {
  const name = String(family || "");
  if (/georgia/i.test(name)) return "Georgia";
  if (/mono|courier/i.test(name)) return "Courier New";
  return "Arial";
}

// Drehwinkel in Grad, 0..360 (im Uhrzeigersinn, wie CSS)
function normalizedRotation(degrees) {
  const value = ((Number(degrees) || 0) % 360 + 360) % 360;
  return value < 0.01 || value > 359.99 ? 0 : value;
}

function pptxFamily(family) {
  return esc(plainFamily(family));
}

// ---------- PPTX schreiben ----------

function pptxRun(text, box) {
  const sz = Math.round(clamp(box.fontPt, 1, 400) * 100);
  const spc = box.spacingPt ? ` spc="${Math.round(box.spacingPt * 100)}"` : "";
  const rPr = `<a:rPr lang="de-DE" sz="${sz}" b="${box.bold ? 1 : 0}" i="${box.italic ? 1 : 0}"${spc} dirty="0">${pptxColor(box.color, box.alpha)}<a:latin typeface="${pptxFamily(box.family)}"/><a:cs typeface="${pptxFamily(box.family)}"/></a:rPr>`;
  return `<a:r>${rPr}<a:t>${esc(text)}</a:t></a:r>`;
}

function pptxParagraph(line, box) {
  const algn = box.align === "ctr" ? "ctr" : box.align === "r" ? "r" : box.align === "just" ? "just" : "l";
  const ln = box.lineSpacingPt
    ? `<a:lnSpc><a:spcPts val="${Math.round(box.lineSpacingPt * 100)}"/></a:lnSpc>`
    : "";
  const end = `<a:endParaRPr lang="de-DE" sz="${Math.round(clamp(box.fontPt, 1, 400) * 100)}"/>`;
  const body = line === "" ? "" : pptxRun(line, box);
  return `<a:p><a:pPr algn="${algn}">${ln}</a:pPr>${body}${end}</a:p>`;
}

function pptxTextShape(box, id) {
  const x = Math.round(box.x * PPTX_W);
  const y = Math.round(box.y * PPTX_H);
  const cx = Math.max(1, Math.round(box.w * PPTX_W));
  const cy = Math.max(1, Math.round(box.h * PPTX_H));
  const wrap = box.wrap === false ? "none" : "square";
  const rot = normalizedRotation(box.rot) ? ` rot="${Math.round(normalizedRotation(box.rot) * 60000)}"` : "";
  const paragraphs = box.text.split("\n").map((line) => pptxParagraph(line, box)).join("");
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${rot}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="${wrap}" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

function pptxPicture(relId, id) {
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Folienbild ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${PPTX_W}" cy="${PPTX_H}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

const PPTX_GROUP_PROPS = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

function pptxSlideXml(slide, mediaRelId) {
  const bg = slide.bg
    ? `<p:bg><p:bgPr>${pptxColor(slide.bg)}<a:effectLst/></p:bgPr></p:bg>`
    : "";
  let nextId = 2;
  const parts = [];
  if (mediaRelId) parts.push(pptxPicture(mediaRelId, nextId++));
  slide.texts.forEach((box) => parts.push(pptxTextShape(box, nextId++)));
  return `${XML_HEADER}<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld>${bg}<p:spTree>${PPTX_GROUP_PROPS}${parts.join("")}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

const PPTX_THEME = `${XML_HEADER}<a:theme xmlns:a="${NS.a}" name="GFS Studio"><a:themeElements><a:clrScheme name="GFS"><a:dk1><a:srgbClr val="171717"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="57554F"/></a:dk2><a:lt2><a:srgbClr val="EEECE5"/></a:lt2><a:accent1><a:srgbClr val="EF3828"/></a:accent1><a:accent2><a:srgbClr val="4D86A2"/></a:accent2><a:accent3><a:srgbClr val="64815C"/></a:accent3><a:accent4><a:srgbClr val="D7D4CA"/></a:accent4><a:accent5><a:srgbClr val="85837D"/></a:accent5><a:accent6><a:srgbClr val="C3C0B9"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="GFS"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="GFS"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;

const PPTX_MASTER = `${XML_HEADER}<p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${PPTX_GROUP_PROPS}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`;

const PPTX_LAYOUT = `${XML_HEADER}<p:sldLayout xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" type="blank" preserve="1"><p:cSld name="Leer"><p:spTree>${PPTX_GROUP_PROPS}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const REL_HEADER = `${XML_HEADER}<Relationships xmlns="${NS.rel}">`;

function relationship(id, type, target) {
  return `<Relationship Id="${id}" Type="${NS.pkgRel}/${type}" Target="${target}"/>`;
}

function mimeExtension(mime) {
  if (mime === "image/jpeg") return "jpeg";
  if (mime === "image/webp") return "webp";
  return "png";
}

export async function buildPptx(slides) {
  if (!slides.length) throw new Error("Keine Folien zum Exportieren.");
  const zip = new JSZip();
  const contentTypeOverrides = [];
  const presentationRels = [
    relationship("rId1", "slideMaster", "slideMasters/slideMaster1.xml"),
    relationship("rId2", "presProps", "presProps.xml"),
    relationship("rId3", "viewProps", "viewProps.xml"),
    relationship("rId4", "theme", "theme/theme1.xml"),
    relationship("rId5", "tableStyles", "tableStyles.xml"),
  ];
  const slideIds = [];

  slides.forEach((slide, index) => {
    const number = index + 1;
    let mediaRelId = null;
    const slideRels = [relationship("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")];
    if (slide.layer) {
      const ext = mimeExtension(slide.layer.mime);
      zip.file(`ppt/media/slide${number}.${ext}`, slide.layer.bytes, { createFolders: false });
      slideRels.push(relationship("rId2", "image", `../media/slide${number}.${ext}`));
      mediaRelId = "rId2";
    }
    zip.file(`ppt/slides/slide${number}.xml`, pptxSlideXml(slide, mediaRelId));
    zip.file(`ppt/slides/_rels/slide${number}.xml.rels`, `${REL_HEADER}${slideRels.join("")}</Relationships>`);
    contentTypeOverrides.push(`<Override PartName="/ppt/slides/slide${number}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`);
    presentationRels.push(relationship(`rId${100 + number}`, "slide", `slides/slide${number}.xml`));
    slideIds.push(`<p:sldId id="${255 + number}" r:id="rId${100 + number}"/>`);
  });

  const mediaTypes = new Set(slides.filter((s) => s.layer).map((s) => mimeExtension(s.layer.mime)));
  const defaults = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    ...[...mediaTypes].map((ext) => `<Default Extension="${ext}" ContentType="image/${ext === "jpeg" ? "jpeg" : ext}"/>`),
  ];

  zip.file("[Content_Types].xml", `${XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults.join("")}<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>${contentTypeOverrides.join("")}</Types>`);

  zip.file("_rels/.rels", `${REL_HEADER}${relationship("rId1", "officeDocument", "ppt/presentation.xml")}<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>${relationship("rId3", "extended-properties", "docProps/app.xml")}</Relationships>`);
  zip.file("docProps/core.xml", `${XML_HEADER}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>GFS Studio Präsentation</dc:title><dc:creator>GFS Studio</dc:creator><dc:language>de-DE</dc:language><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created></cp:coreProperties>`);
  zip.file("docProps/app.xml", `${XML_HEADER}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>GFS Studio</Application><Slides>${slides.length}</Slides></Properties>`);
  zip.file("ppt/presentation.xml", `${XML_HEADER}<p:presentation xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slideIds.join("")}</p:sldIdLst><p:sldSz cx="${PPTX_W}" cy="${PPTX_H}"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="de-DE"/></a:defPPr></p:defaultTextStyle></p:presentation>`);
  zip.file("ppt/_rels/presentation.xml.rels", `${REL_HEADER}${presentationRels.join("")}</Relationships>`);
  zip.file("ppt/presProps.xml", `${XML_HEADER}<p:presentationPr xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"/>`);
  zip.file("ppt/viewProps.xml", `${XML_HEADER}<p:viewPr xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`);
  zip.file("ppt/tableStyles.xml", `${XML_HEADER}<a:tblStyleLst xmlns:a="${NS.a}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
  zip.file("ppt/slideMasters/slideMaster1.xml", PPTX_MASTER);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", `${REL_HEADER}${relationship("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")}${relationship("rId2", "theme", "../theme/theme1.xml")}</Relationships>`);
  zip.file("ppt/slideLayouts/slideLayout1.xml", PPTX_LAYOUT);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", `${REL_HEADER}${relationship("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")}</Relationships>`);
  zip.file("ppt/theme/theme1.xml", PPTX_THEME);

  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
}

// ---------- ODP schreiben ----------

function odfLength(fraction, total) {
  return `${round(fraction * total, 3)}cm`;
}

function odfText(box, styleIndex) {
  const align = box.align === "ctr" ? "center" : box.align === "r" ? "end" : box.align === "just" ? "justify" : "start";
  const lineHeight = box.lineSpacingPt ? `fo:line-height="${round(box.lineSpacingPt, 2)}pt"` : "";
  const paragraphStyle = `P${styleIndex}`;
  const textStyle = `T${styleIndex}`;
  const textProps = [
    `fo:font-family="${pptxFamily(box.family)}"`,
    `style:font-family-generic="swiss"`,
    `fo:font-size="${round(box.fontPt, 2)}pt"`,
    `fo:font-weight="${box.bold ? "bold" : "normal"}"`,
    `fo:font-style="${box.italic ? "italic" : "normal"}"`,
    `fo:color="${hexOf(box.color).toLowerCase()}"`,
    box.spacingPt ? `fo:letter-spacing="${round(box.spacingPt, 2)}pt"` : "",
  ].filter(Boolean).join(" ");
  const styles = `<style:style style:name="${paragraphStyle}" style:family="paragraph"><style:paragraph-properties fo:text-align="${align}" ${lineHeight}/></style:style><style:style style:name="${textStyle}" style:family="text"><style:text-properties ${textProps}/></style:style>`;
  const paragraphs = box.text.split("\n").map((line) => {
    const span = line === "" ? "" : `<text:span text:style-name="${textStyle}">${esc(line)}</text:span>`;
    return `<text:p text:style-name="${paragraphStyle}">${span}</text:p>`;
  }).join("");
  return { styles, paragraphs };
}

export async function buildOdp(slides, { template = false } = {}) {
  if (!slides.length) throw new Error("Keine Folien zum Exportieren.");
  const mimetype = template
    ? "application/vnd.oasis.opendocument.presentation-template"
    : "application/vnd.oasis.opendocument.presentation";
  const zip = new JSZip();
  // mimetype muss die erste, unkomprimierte Datei des Pakets sein
  zip.file("mimetype", mimetype, { compression: "STORE", createFolders: false });
  const autoStyles = [];
  const pages = [];
  const manifestEntries = [];
  let styleIndex = 0;

  slides.forEach((slide, pageIndex) => {
    const number = pageIndex + 1;
    const bg = slide.bg ? hexOf(slide.bg).toLowerCase() : "#ffffff";
    autoStyles.push(`<style:style style:name="dp${number}" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="${bg}" draw:background-size="border"/></style:style>`);
    const shapes = [];
    if (slide.layer) {
      const ext = mimeExtension(slide.layer.mime);
      const href = `Pictures/slide${number}.${ext}`;
      zip.file(href, slide.layer.bytes, { createFolders: false });
      manifestEntries.push(`<manifest:file-entry manifest:full-path="${href}" manifest:media-type="${slide.layer.mime}"/>`);
      autoStyles.push(`<style:style style:name="gimg${number}" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none"/></style:style>`);
      shapes.push(`<draw:frame draw:style-name="gimg${number}" draw:name="Folienbild ${number}" svg:x="0cm" svg:y="0cm" svg:width="${ODP_W_CM}cm" svg:height="${ODP_H_CM}cm" draw:z-index="0"><draw:image xlink:href="${href}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`);
    }
    slide.texts.forEach((box, index) => {
      styleIndex += 1;
      const { styles, paragraphs } = odfText(box, styleIndex);
      autoStyles.push(styles);
      const frameStyle = box.wrap === false ? "gtextn" : "gtext";
      const angle = normalizedRotation(box.rot);
      const transform = angle ? ` draw:transform="rotate (${round((angle * Math.PI) / 180, 5)})"` : "";
      shapes.push(`<draw:frame draw:style-name="${frameStyle}" draw:name="Text ${number}-${index + 1}"${transform} svg:x="${odfLength(box.x, ODP_W_CM)}" svg:y="${odfLength(box.y, ODP_H_CM)}" svg:width="${odfLength(box.w, ODP_W_CM)}" svg:height="${odfLength(box.h, ODP_H_CM)}"><draw:text-box>${paragraphs}</draw:text-box></draw:frame>`);
    });
    pages.push(`<draw:page draw:name="Folie ${number}" draw:style-name="dp${number}" draw:master-page-name="Default">${shapes.join("")}</draw:page>`);
  });

  autoStyles.push('<style:style style:name="gtext" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none" fo:padding="0cm" fo:border="none" draw:textarea-horizontal-align="justify" draw:textarea-vertical-align="top" draw:auto-grow-height="false" draw:auto-grow-width="false" fo:wrap-option="wrap"/></style:style>');
  autoStyles.push('<style:style style:name="gtextn" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none" fo:padding="0cm" fo:border="none" draw:textarea-horizontal-align="justify" draw:textarea-vertical-align="top" draw:auto-grow-height="false" draw:auto-grow-width="false" fo:wrap-option="no-wrap"/></style:style>');

  const ns = `xmlns:office="${ODF.office}" xmlns:style="${ODF.style}" xmlns:text="${ODF.text}" xmlns:draw="${ODF.draw}" xmlns:fo="${ODF.fo}" xmlns:svg="${ODF.svg}" xmlns:xlink="${ODF.xlink}" xmlns:presentation="${ODF.presentation}"`;
  const content = `${XML_HEADER}<office:document-content ${ns} office:version="1.3"><office:font-face-decls><style:font-face style:name="Arial" svg:font-family="Arial" style:font-family-generic="swiss"/><style:font-face style:name="Georgia" svg:font-family="Georgia" style:font-family-generic="roman"/><style:font-face style:name="Courier New" svg:font-family="'Courier New'" style:font-family-generic="modern"/></office:font-face-decls><office:automatic-styles>${autoStyles.join("")}</office:automatic-styles><office:body><office:presentation>${pages.join("")}</office:presentation></office:body></office:document-content>`;
  const styles = `${XML_HEADER}<office:document-styles ${ns} office:version="1.3"><office:styles/><office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="${ODP_W_CM}cm" fo:page-height="${ODP_H_CM}cm" style:print-orientation="landscape" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles></office:document-styles>`;
  const meta = `${XML_HEADER}<office:document-meta xmlns:office="${ODF.office}" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><dc:title>GFS Studio Präsentation</dc:title><dc:language>de-DE</dc:language><meta:generator>GFS Studio</meta:generator></office:meta></office:document-meta>`;
  const manifest = `${XML_HEADER}<manifest:manifest xmlns:manifest="${ODF.manifest}" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${mimetype}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>${manifestEntries.join("")}</manifest:manifest>`;

  zip.file("content.xml", content);
  zip.file("styles.xml", styles);
  zip.file("meta.xml", meta);
  zip.file("META-INF/manifest.xml", manifest, { createFolders: false });

  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", mimeType: mimetype });
}

// ---------- Lesen (gemeinsame Helfer) ----------

function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("Die Datei enthält beschädigtes XML.");
  return doc;
}

function byTag(root, ns, name) {
  return Array.from(root.getElementsByTagNameNS(ns, name));
}

function directChildren(el, ns, name) {
  return Array.from(el.children).filter((child) => child.localName === name && child.namespaceURI === ns);
}

function parseColor(hex) {
  return /^[0-9a-f]{6}$/i.test(hex || "") ? `#${hex.toLowerCase()}` : null;
}

function familyFromTypeface(typeface = "") {
  const name = typeface.toLowerCase();
  if (/georgia|times|serif/.test(name) && !/sans/.test(name)) return "Georgia, serif";
  if (/courier|mono|consol/.test(name)) return "monospace";
  return "Arial, Helvetica, sans-serif";
}

function alignFromWord(word) {
  if (word === "ctr" || word === "center") return "ctr";
  if (word === "r" || word === "end" || word === "right") return "r";
  if (word === "just" || word === "justify") return "just";
  return "l";
}

const MIME_BY_EXT = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml", emf: "image/emf", wmf: "image/wmf", tif: "image/tiff", tiff: "image/tiff" };

function mimeFromPath(path) {
  const ext = (path.split(".").pop() || "").toLowerCase();
  return MIME_BY_EXT[ext] || "application/octet-stream";
}

function resolveZipPath(baseDir, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = `${baseDir}/${target}`.split("/");
  const out = [];
  parts.forEach((part) => {
    if (part === "" || part === ".") return;
    if (part === "..") out.pop();
    else out.push(part);
  });
  return out.join("/");
}

// ---------- PPTX lesen ----------

async function readRels(zip, path) {
  const file = zip.file(path);
  if (!file) return new Map();
  const doc = parseXml(await file.async("text"));
  const map = new Map();
  byTag(doc, NS.rel, "Relationship").forEach((rel) => map.set(rel.getAttribute("Id"), rel.getAttribute("Target")));
  return map;
}

function pptxTextFromBody(txBody) {
  const paragraphs = byTag(txBody, NS.a, "p").map((p) => {
    let line = "";
    Array.from(p.getElementsByTagNameNS(NS.a, "*")).forEach((node) => {
      if (node.localName === "t") line += node.textContent;
      if (node.localName === "br") line += "\n";
    });
    return line;
  });
  return paragraphs.join("\n").replace(/\n+$/, "");
}

function pptxRunStyle(txBody) {
  const style = { fontPt: 18, color: "#171717", bold: false, italic: false, family: "Arial, Helvetica, sans-serif", align: "l" };
  const firstRun = byTag(txBody, NS.a, "rPr")[0];
  if (firstRun) {
    const sz = Number(firstRun.getAttribute("sz"));
    if (sz) style.fontPt = sz / 100;
    style.bold = firstRun.getAttribute("b") === "1";
    style.italic = firstRun.getAttribute("i") === "1";
    const fill = byTag(firstRun, NS.a, "srgbClr")[0];
    if (fill) style.color = parseColor(fill.getAttribute("val")) || style.color;
    const latin = byTag(firstRun, NS.a, "latin")[0];
    if (latin) style.family = familyFromTypeface(latin.getAttribute("typeface"));
  }
  const firstPara = byTag(txBody, NS.a, "pPr")[0];
  if (firstPara) style.align = alignFromWord(firstPara.getAttribute("algn"));
  return style;
}

function pptxFrame(xfrm, slideW, slideH) {
  const off = byTag(xfrm, NS.a, "off")[0];
  const ext = byTag(xfrm, NS.a, "ext")[0];
  if (!off || !ext) return null;
  return {
    x: Number(off.getAttribute("x")) / slideW,
    y: Number(off.getAttribute("y")) / slideH,
    w: Number(ext.getAttribute("cx")) / slideW,
    h: Number(ext.getAttribute("cy")) / slideH,
  };
}

export async function parsePptx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const presentation = zip.file("ppt/presentation.xml");
  if (!presentation) throw new Error("Die Datei ist keine gültige PPTX-Präsentation.");
  const presDoc = parseXml(await presentation.async("text"));
  const sldSz = byTag(presDoc, NS.p, "sldSz")[0];
  const slideW = Number(sldSz?.getAttribute("cx")) || PPTX_W;
  const slideH = Number(sldSz?.getAttribute("cy")) || PPTX_H;
  const presRels = await readRels(zip, "ppt/_rels/presentation.xml.rels");
  const order = byTag(presDoc, NS.p, "sldId").map((node) => presRels.get(node.getAttributeNS(NS.r, "id")));
  const slidePaths = order.filter(Boolean).map((target) => resolveZipPath("ppt", target));
  if (!slidePaths.length) throw new Error("Keine Folien in der PPTX-Datei gefunden.");

  const slides = [];
  for (const slidePath of slidePaths) {
    const file = zip.file(slidePath);
    if (!file) continue;
    const doc = parseXml(await file.async("text"));
    const dir = slidePath.split("/").slice(0, -1).join("/");
    const rels = await readRels(zip, `${dir}/_rels/${slidePath.split("/").pop()}.rels`);
    const bgColor = byTag(doc, NS.p, "bg")[0];
    const bg = bgColor ? parseColor(byTag(bgColor, NS.a, "srgbClr")[0]?.getAttribute("val")) : null;
    const elements = [];
    const tree = byTag(doc, NS.p, "spTree")[0];
    if (tree) {
      for (const node of Array.from(tree.children)) {
        if (node.localName === "sp") {
          const spPr = directChildren(node, NS.p, "spPr")[0];
          const xfrm = spPr && byTag(spPr, NS.a, "xfrm")[0];
          const frame = xfrm && pptxFrame(xfrm, slideW, slideH);
          if (!frame) continue;
          const txBody = directChildren(node, NS.p, "txBody")[0];
          const text = txBody ? pptxTextFromBody(txBody) : "";
          if (text.trim()) {
            elements.push({ type: "text", text, ...frame, ...pptxRunStyle(txBody) });
            continue;
          }
          const geom = byTag(spPr, NS.a, "prstGeom")[0];
          const prst = geom?.getAttribute("prst") || "rect";
          const fillEl = directChildren(spPr, NS.a, "solidFill")[0];
          const color = parseColor(fillEl && byTag(fillEl, NS.a, "srgbClr")[0]?.getAttribute("val"));
          if (!color) continue;
          const shape = prst === "ellipse" ? "circle" : prst === "line" || prst === "straightConnector1" ? "line" : "rectangle";
          elements.push({ type: "shape", shape, ...frame, color });
        } else if (node.localName === "pic") {
          const spPr = directChildren(node, NS.p, "spPr")[0];
          const xfrm = spPr && byTag(spPr, NS.a, "xfrm")[0];
          const frame = xfrm && pptxFrame(xfrm, slideW, slideH);
          const blip = byTag(node, NS.a, "blip")[0];
          const embed = blip?.getAttributeNS(NS.r, "embed");
          const target = embed && rels.get(embed);
          const mediaFile = target && zip.file(resolveZipPath(dir, target));
          if (!frame || !mediaFile) continue;
          const mediaPath = resolveZipPath(dir, target);
          elements.push({ type: "image", ...frame, mime: mimeFromPath(mediaPath), bytes: await mediaFile.async("uint8array") });
        }
      }
    }
    slides.push({ bg, elements });
  }
  return { slides };
}

// ---------- ODP lesen ----------

const UNIT_TO_CM = { cm: 1, mm: 0.1, in: 2.54, inch: 2.54, pt: 2.54 / 72, pc: 2.54 / 6, px: 2.54 / 96 };

function lengthToCm(value) {
  const match = String(value || "").trim().match(/^(-?[\d.]+)\s*([a-z]*)$/i);
  if (!match) return 0;
  return Number(match[1]) * (UNIT_TO_CM[match[2].toLowerCase()] ?? 1);
}

function odfTextContent(node) {
  let out = "";
  node.childNodes.forEach((child) => {
    if (child.nodeType === 3) {
      out += child.textContent;
    } else if (child.nodeType === 1) {
      const name = child.localName;
      if (name === "s") out += " ".repeat(Number(child.getAttribute("text:c")) || 1);
      else if (name === "line-break") out += "\n";
      else if (name === "tab") out += "\t";
      else out += odfTextContent(child);
    }
  });
  return out;
}

function collectOdfStyles(root, map) {
  byTag(root, ODF.style, "style").forEach((style) => {
    const name = style.getAttributeNS(ODF.style, "name");
    if (!name) return;
    const entry = map.get(name) || {};
    const gfx = byTag(style, ODF.style, "graphic-properties")[0];
    if (gfx) {
      const fill = gfx.getAttributeNS(ODF.draw, "fill");
      const fillColor = gfx.getAttributeNS(ODF.draw, "fill-color");
      if (fill === "solid" && parseColor(fillColor?.replace("#", ""))) entry.fill = fillColor.toLowerCase();
    }
    const draw = byTag(style, ODF.style, "drawing-page-properties")[0];
    if (draw) {
      const fillColor = draw.getAttributeNS(ODF.draw, "fill-color");
      if (fillColor) entry.pageFill = fillColor.toLowerCase();
    }
    const para = byTag(style, ODF.style, "paragraph-properties")[0];
    if (para) entry.align = alignFromWord(para.getAttributeNS(ODF.fo, "text-align"));
    const text = byTag(style, ODF.style, "text-properties")[0];
    if (text) {
      const size = text.getAttributeNS(ODF.fo, "font-size");
      if (size) entry.fontPt = lengthToCm(size) / (2.54 / 72);
      entry.bold = text.getAttributeNS(ODF.fo, "font-weight") === "bold";
      entry.italic = text.getAttributeNS(ODF.fo, "font-style") === "italic";
      const color = text.getAttributeNS(ODF.fo, "color");
      if (color) entry.color = color.toLowerCase();
      const family = text.getAttributeNS(ODF.fo, "font-family");
      if (family) entry.family = familyFromTypeface(family.replace(/['"]/g, ""));
    }
    map.set(name, entry);
  });
  return map;
}

export async function parseOdp(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const contentFile = zip.file("content.xml");
  if (!contentFile) throw new Error("Keine content.xml in der ODP-Datei gefunden.");
  const contentDoc = parseXml(await contentFile.async("text"));
  const stylesFile = zip.file("styles.xml");
  const stylesDoc = stylesFile ? parseXml(await stylesFile.async("text")) : null;

  const styleMap = new Map();
  collectOdfStyles(contentDoc, styleMap);
  if (stylesDoc) collectOdfStyles(stylesDoc, styleMap);

  let pageWidthCm = ODP_W_CM;
  let pageHeightCm = ODP_H_CM;
  if (stylesDoc) {
    const layout = byTag(stylesDoc, ODF.style, "page-layout-properties")[0];
    if (layout) {
      pageWidthCm = lengthToCm(layout.getAttributeNS(ODF.fo, "page-width")) || pageWidthCm;
      pageHeightCm = lengthToCm(layout.getAttributeNS(ODF.fo, "page-height")) || pageHeightCm;
    }
  }

  const pages = byTag(contentDoc, ODF.draw, "page");
  if (!pages.length) throw new Error("Keine Folien in der ODP-Datei gefunden.");

  const slides = [];
  for (const page of pages) {
    const pageStyle = styleMap.get(page.getAttributeNS(ODF.draw, "style-name")) || {};
    const elements = [];
    const frameOf = (node) => ({
      x: lengthToCm(node.getAttributeNS(ODF.svg, "x")) / pageWidthCm,
      y: lengthToCm(node.getAttributeNS(ODF.svg, "y")) / pageHeightCm,
      w: lengthToCm(node.getAttributeNS(ODF.svg, "width")) / pageWidthCm,
      h: lengthToCm(node.getAttributeNS(ODF.svg, "height")) / pageHeightCm,
    });
    for (const node of Array.from(page.children)) {
      const style = styleMap.get(node.getAttributeNS(ODF.draw, "style-name")) || {};
      const hasDirectText = directChildren(node, ODF.text, "p").length > 0 || directChildren(node, ODF.text, "h").length > 0;
      if (node.localName === "frame" || node.localName === "custom-shape" || node.localName === "rect" || node.localName === "ellipse" || node.localName === "text-box") {
        const textBox = node.localName === "frame" ? directChildren(node, ODF.draw, "text-box")[0] : (hasDirectText ? node : null);
        const image = node.localName === "frame" ? directChildren(node, ODF.draw, "image")[0] : null;
        if (textBox) {
          const paragraphs = directChildren(textBox, ODF.text, "p").concat(directChildren(textBox, ODF.text, "h"));
          const lines = paragraphs.map((p) => odfTextContent(p));
          const text = lines.join("\n").replace(/\n+$/, "");
          if (!text.trim()) continue;
          const span = byTag(textBox, ODF.text, "span")[0];
          const spanStyle = span ? styleMap.get(span.getAttributeNS(ODF.text, "style-name")) || {} : {};
          const paraStyle = styleMap.get(paragraphs[0]?.getAttributeNS(ODF.text, "style-name")) || {};
          elements.push({
            type: "text",
            text,
            ...frameOf(node),
            fontPt: spanStyle.fontPt || style.fontPt || 18,
            color: spanStyle.color ? parseColor(spanStyle.color.replace("#", "")) || "#171717" : "#171717",
            bold: Boolean(spanStyle.bold),
            italic: Boolean(spanStyle.italic),
            family: spanStyle.family || "Arial, Helvetica, sans-serif",
            align: paraStyle.align || "l",
          });
        } else if (image) {
          const href = image.getAttributeNS(ODF.xlink, "href");
          const file = href && zip.file(href.replace(/^\.\//, ""));
          if (!file) continue;
          elements.push({ type: "image", ...frameOf(node), mime: mimeFromPath(href), bytes: await file.async("uint8array") });
        }
      } else if (node.localName === "rect" || node.localName === "ellipse" || node.localName === "custom-shape") {
        const color = style.fill ? parseColor(style.fill.replace("#", "")) : null;
        if (!color) continue;
        elements.push({ type: "shape", shape: node.localName === "ellipse" ? "circle" : "rectangle", ...frameOf(node), color });
      }
    }
    slides.push({ bg: pageStyle.pageFill ? parseColor(pageStyle.pageFill.replace("#", "")) : null, elements });
  }
  return { slides };
}
