import os
import zipfile

from procurar import ler_chaves_do_texto, zipar_pasta


def test_ler_chaves_do_texto_ignora_duplicadas():
    texto = "12345678901234567890123456789012345678901234\n12345678901234567890123456789012345678901234\n"

    assert ler_chaves_do_texto(texto) == ["12345678901234567890123456789012345678901234"]


def test_criar_zip_backup_gera_arquivo(tmp_path):
    xml = tmp_path / "arquivo_12345678901234567890123456789012345678901234.xml"
    xml.write_text("<xml>teste</xml>", encoding="utf-8")

    resultado = zipar_pasta(str(tmp_path), somente_xml=True)

    assert resultado["sucesso"] is True
    zip_path = resultado["zip"]
    assert os.path.exists(zip_path)
    with zipfile.ZipFile(zip_path, "r") as zf:
        assert xml.name in zf.namelist()


def test_limita_payload_resend_por_anexo(tmp_path):
    arquivo = tmp_path / "teste.xml"
    arquivo.write_text("<xml>ok</xml>", encoding="utf-8")

    resultado = zipar_pasta(str(tmp_path), somente_xml=True)

    assert resultado["sucesso"] is True
    assert resultado["zip"].lower().endswith(".zip")
