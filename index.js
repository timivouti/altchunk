var Docxtemplater = require('docxtemplater');
var JSZip = require('jszip');

export function mergeDocuments(firstDocument, files, isDraft, error) {
  try {
    var watermark = '<w:sdt><w:sdtPr><w:id w:val="-1960019101"/><w:docPartObj><w:docPartGallery w:val="Watermarks"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:r><w:pict w14:anchorId="59CEE41C"><v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype><v:shape id="PowerPlusWaterMarkObject357831064" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:412.4pt;height:247.45pt;rotation:315;z-index:-251656704;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;calibri&quot;;font-size:1pt" string="LUONNOS"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r></w:sdtContent></w:sdt>';

    String.prototype.splice = function (idx, rem, str) {
      return this.slice(0, idx) + str + this.slice(idx + Math.abs(rem));
    };

    function utf8ArrayToString(array) {
      var out, i, len, c;
      var char2, char3;

      out = "";
      len = array.length;
      i = 0;
      while (i < len) {
        c = array[i++];
        switch (c >> 4) {
          case 0: case 1: case 2: case 3: case 4: case 5: case 6: case 7:
            // 0xxxxxxx
            out += String.fromCharCode(c);
            break;
          case 12: case 13:
            // 110x xxxx   10xx xxxx
            char2 = array[i++];
            out += String.fromCharCode(((c & 0x1F) << 6) | (char2 & 0x3F));
            break;
          case 14:
            // 1110 xxxx  10xx xxxx  10xx xxxx
            char2 = array[i++];
            char3 = array[i++];
            out += String.fromCharCode(((c & 0x0F) << 12) |
              ((char2 & 0x3F) << 6) |
              ((char3 & 0x3F) << 0));
            break;
        }
      }

      return out;
    }

    var zip = new JSZip(firstDocument);

    var document = utf8ArrayToString(zip.file("word/document.xml")._data.getContent());
    var newDocument = document;

    var documentXmlRels = utf8ArrayToString(zip.file("word/_rels/document.xml.rels")._data.getContent());
    var newDocumentXmlRels = documentXmlRels;

    var contentType = utf8ArrayToString(zip.file("[Content_Types].xml")._data.getContent());
    var newContentType = contentType;
    if (contentType.indexOf('<Default ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" Extension="docx"/>') === -1) {
      newContentType = newContentType.splice(contentType.indexOf("</Types>"), 0, '<Default ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" Extension="docx"/>');
    }

    zip.file("[Content_Types].xml", newContentType.replace(' standalone="true"', ""));

    var notNull = true;
    var howManyAltChunksFound = 0;
    var loop = 1;

    do {
      if (zip.file("word/afchunk" + loop + ".docx") != null) {
        loop++;
        howManyAltChunksFound++;
      } else {
        notNull = false;
      }
    } while (notNull);

    if (isDraft && zip.file("word/header3.xml") != null) {
      var addHeader = utf8ArrayToString(zip.file("word/header3.xml")._data.getContent());
      if (addHeader.indexOf(watermark) === -1) {
        var newHeader = addHeader.splice(addHeader.indexOf("</w:pPr>") + 8, 0, watermark);
        zip.file("word/header3.xml", newHeader);
      }
    }

    if (zip.file("docProps/custom.xml") != null) {
      var addCustomXml = utf8ArrayToString(zip.file("docProps/custom.xml")._data.getContent());
      var count = (addCustomXml.match(/pid/g) || []).length + 2;
      addCustomXml = addCustomXml.replace("0x0101008FF0A72AFE67D442B57B8AACC253BDB7005D0E72A7CB9D6E458F65F43B2D3DBB6D", "0x010100C8CB6112EF7BD440B47A9D66E3EE21B00043AB5F92ADB0874286AA9CE9D68ADE37");
      if (addCustomXml.indexOf("Kera_Dokumentin tila") > -1) {
        if (addCustomXml.indexOf("Luonnos") > -1) {
          zip.file("docProps/custom.xml", isDraft ? addCustomXml : addCustomXml.replace("Luonnos", "Julkaisu"));
        }
        if (addCustomXml.indexOf("Julkaisu") > -1) {
          zip.file("docProps/custom.xml", isDraft ? addCustomXml.replace("Julkaisu", "Luonnos") : addCustomXml);
        }
      } else {
        var newCustomXml = addCustomXml.splice(addCustomXml.indexOf("</Properties>"), 0, '<property name="Kera_Dokumentin tila" pid="' + count + ' fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"><vt:lpwstr>' + isDraft ? "Luonnos" : "Julkaisu" + '</vt:lpwstr></property>');
        zip.file("docProps/custom.xml", newCustomXml);
      }
    }

    for (var i = 0; i < files.length; i++) {
      var altChunkId = howManyAltChunksFound + (i + 1);
      zip.file('word/afchunk' + altChunkId + '.docx', files[i], { binary: true });
      newDocument = newDocument.splice(newDocument.lastIndexOf("<w:sectPr"), 0, '<w:altChunk r:id="AltChunkId' + altChunkId + '"/>');

      newDocumentXmlRels = newDocumentXmlRels.splice(newDocumentXmlRels.indexOf("</Relationships>"), 0, '<Relationship Id="AltChunkId' + altChunkId + '" Target="/word/afchunk' + altChunkId + '.docx" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"/>');
    }

    zip.file("word/document.xml", newDocument);
    zip.file("word/_rels/document.xml.rels", newDocumentXmlRels);

    var doc = new Docxtemplater().loadZip(zip);

    doc.render();

    var out = doc.getZip().generate({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });
    return out;

  } catch (error) {
    console.log(error);
    return fileError;
  }
}
