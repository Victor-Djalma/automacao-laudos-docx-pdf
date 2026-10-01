from __future__ import annotations

import sys
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from automacao_laudos import (  # noqa: E402
    W_NS,
    document_text,
    number_documents,
    read_document_xml,
)


def make_docx(path: Path, date: str, title: str = "Relatório de Ensaio N°") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    document = etree.Element(f"{{{W_NS}}}document", nsmap={"w": W_NS})
    body = etree.SubElement(document, f"{{{W_NS}}}body")
    for value in [title, f"Data do Recebimento: {date}", "Conteúdo preservado"]:
        paragraph = etree.SubElement(body, f"{{{W_NS}}}p")
        run = etree.SubElement(paragraph, f"{{{W_NS}}}r")
        text = etree.SubElement(run, f"{{{W_NS}}}t")
        text.text = value
    xml = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", xml)
        archive.writestr("preservado.bin", b"nao alterar")


def test_numbers_in_date_order_and_preserves_other_parts(tmp_path: Path) -> None:
    source = tmp_path / "entrada"
    output = tmp_path / "saida"
    make_docx(source / "B" / "segundo.docx", "11/08/2026")
    make_docx(source / "A" / "primeiro.docx", "03/08/2026")
    rows, excluded = number_documents(source, output, 1072, 2026, False)
    assert excluded == []
    assert [row["report_number"] for row in rows] == ["01072.2026", "01073.2026"]
    assert rows[0]["relative_path"] == "A/primeiro.docx"
    text = document_text(read_document_xml(output / "A" / "primeiro.docx"))
    assert "Relatório de Ensaio N° 01072.2026" in text
    with zipfile.ZipFile(output / "A" / "primeiro.docx") as archive:
        assert archive.read("preservado.bin") == b"nao alterar"


def test_replaces_existing_number_and_marks_old_year(tmp_path: Path) -> None:
    source = tmp_path / "entrada"
    output = tmp_path / "saida"
    make_docx(source / "laudo.docx", "03/08/2025", "Relatório de Ensaio N° 99999.2024")
    rows, _ = number_documents(source, output, 1, 2026, False)
    assert rows[0]["outside_report_year"] is True
    text = document_text(read_document_xml(output / "laudo.docx"))
    assert "00001.2026" in text
    assert "99999.2024" not in text

