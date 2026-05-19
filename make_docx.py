"""
Gumagawa ng valid .docx file gamit ang Python standard library lamang.
Ang .docx ay isang ZIP archive na naglalaman ng XML files (Office Open XML format).
"""

import zipfile
from xml.sax.saxutils import escape


def make_docx(filename: str, title: str, paragraphs: list[str]) -> None:
    # ---- [Content_Types].xml ----
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    # ---- _rels/.rels ----
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    # ---- word/document.xml ----
    body_parts = []

    # Title (heading)
    body_parts.append(f"""    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="40"/></w:rPr><w:t>{escape(title)}</w:t></w:r>
    </w:p>""")

    # Body paragraphs
    for p in paragraphs:
        body_parts.append(f"""    <w:p>
      <w:r><w:t xml:space="preserve">{escape(p)}</w:t></w:r>
    </w:p>""")

    body_xml = "\n".join(body_parts)

    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
{body_xml}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    # ---- Build the .docx (ZIP) ----
    with zipfile.ZipFile(filename, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)

    print(f"Nagawa: {filename}")


if __name__ == "__main__":
    make_docx(
        "sample.docx",
        title="Halimbawang Dokumento",
        paragraphs=[
            "Ito ay isang sample na .docx file na ginawa ng Kiro.",
            "Pwede mo itong i-edit gamit ang Microsoft Word, Google Docs, o LibreOffice.",
            "Kung gusto mo ng mas detalyadong content, sabihin mo lang sa chat!",
        ],
    )
