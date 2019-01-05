var Docxtemplater = require('docxtemplater');
var JSZip = require('jszip');

export const mergeDocuments = (firstDocument, files, isDraft) => {
  String.prototype.splice = function (idx, rem, str) {
    return this.slice(0, idx) + str + this.slice(idx + Math.abs(rem));
  };

  const utf8ArrayToString = function(array) {
		var out, i, len, c;
		var char2, char3;

		out = "";
		len = array.length;
		i = 0;
		while(i < len) {
		c = array[i++];
		switch(c >> 4)
		{
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

  if (isDraft && zip.file("word/header3.xml") != null) {
    var addHeader = utf8ArrayToString(zip.file("word/header3.xml")._data.getContent());
    var newHeader = addHeader.splice(addHeader.indexOf(`</w:pPr>`) + 8, 0, watermark);
    zip.file("word/header3.xml", newHeader);
  }

  files.map((x, i) => {
    var docxZip = new JSZip(x);
    var docxTemp = new Docxtemplater().loadZip(docxZip);

    docxTemp.render();

    var docxFile = docxTemp.getZip().generate({
          type: "blob",
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });
        
    zip.file(`word/afchunk{i+1}.docx`, docxFile);

    var document = utf8ArrayToString(zip.file("word/document.xml")._data.getContent());
    var newDocument = document.splice(0, document.lastIndexOf("<w:sectPr>"), `<w:altChunk r:id="AltChunkId${i}"/>`);
    zip.file("word/document.xml", newDocument);

    var documentXmlRels = utf8ArrayToString(zip.file("word/_rels/document.xml.rels")._data.getContent());
    var newDocumentXMlRels = documentXmlRels(0, documentXmlRels.lastIndexOf("</Relationships>"), `<Relationship Id="AltChunkId${i}" Target="/word/afchunk${i}.docx" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"/>`);
    zip.file("word/_rels/document.xml.rels", newDocumentXmlRels);
  });

  var doc = new Docxtemplater().loadZip(zip);

  doc.render();

  var out = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      });
  return out;
};
