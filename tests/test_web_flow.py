from __future__ import annotations

import io
import hashlib
import ipaddress
import json
import re

import pandas as pd
import pytest

flask = pytest.importorskip("flask")

from data.models import CrossStepSpec, FileSpec, Recipe

import app as web_app


@pytest.fixture
def client(tmp_path):
    web_app.app.config.update(TESTING=True, SESSION_COOKIE_SECURE=False)
    web_app._audit.directory = tmp_path / "logs"
    web_app._workflows.clear()
    with web_app.app.test_client() as test_client:
        yield test_client
    web_app._workflows.clear()


def _csrf(html: str) -> str:
    match = re.search(r'name="csrf_token" value="([^"]+)"', html)
    assert match
    return match.group(1)


def _excel_bytes(title: str, rows: list[list[str]]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame([[title, None], ["CUIL", "Estado"], *rows]).to_excel(
            writer, header=False, index=False
        )
    return buffer.getvalue()


def test_local_upload_header_selection_cross_preview_and_export(client):
    home = client.get("/")
    token = _csrf(home.get_data(as_text=True))
    left = _excel_bytes("Archivo de altas", [["20123456789", "Activo"], ["20234567890", "Baja"]])
    right = _excel_bytes("Archivo de control", [["20123456789", "OK"], ["20345678901", "OK"]])

    uploaded = client.post(
        "/upload",
        data={
            "csrf_token": token,
            "files": [
                (io.BytesIO(left), "altas.xlsx"),
                (io.BytesIO(right), "control.xlsx"),
            ],
        },
        content_type="multipart/form-data",
    )
    assert uploaded.status_code == 200
    assert "Fila 2" in uploaded.get_data(as_text=True)
    resumed = client.get("/")
    assert "Hay archivos ya cargados" in resumed.get_data(as_text=True)
    assert "Continuar con esta carga" in resumed.get_data(as_text=True)

    headers = client.post(
        "/headers",
        data={"csrf_token": token, "header_0": "1", "header_1": "1"},
    )
    assert headers.status_code == 200
    assert "CUIL" in headers.get_data(as_text=True)
    assert "Limpieza de datos" in headers.get_data(as_text=True)
    assert 'data-show-when="operation:fill_empty" hidden' in headers.get_data(as_text=True)
    previous_step = client.get("/headers")
    assert previous_step.status_code == 200
    assert "altas.xlsx" in previous_step.get_data(as_text=True)
    assert "Cancelar y borrar la carga" not in headers.get_data(as_text=True)

    result = client.post(
        "/run",
        data={
            "csrf_token": token,
            "left_key_0": ["CUIL", "", "", "", ""],
            "right_key_0": ["CUIL", "", "", "", ""],
            "join_type": "inner",
            "normalize": "yes",
        },
    )
    assert result.status_code == 200
    assert "Columnas calculadas y reglas" in result.get_data(as_text=True)
    assert 'data-show-when="rule_type:if_then_else" hidden' in result.get_data(as_text=True)
    cross_back = client.get("/cross-config")
    assert cross_back.status_code == 200
    assert "CUIL" in cross_back.get_data(as_text=True)
    result = client.post(
        "/rules/skip",
        data={"csrf_token": token},
        follow_redirects=True,
    )
    assert result.status_code == 200
    assert "Revisar resultado antes de exportar" in result.get_data(as_text=True)
    assert "20123456789" in result.get_data(as_text=True)

    live_preview = client.get("/result/preview", query_string={"quick_search": "20123456789"})
    assert live_preview.status_code == 200
    preview_payload = live_preview.get_json()
    assert preview_payload["filtered_count"] == 1
    assert preview_payload["rows"][0]["values"]["CUIL"] == "20123456789"

    exported = client.post("/export", data={"csrf_token": token})
    assert exported.status_code == 200
    assert exported.mimetype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert exported.data[:2] == b"PK"

    extension = client.post(
        "/extend",
        data={
            "csrf_token": token,
            "files": (io.BytesIO(_excel_bytes("Premios", [["20123456789", "Pagado"]])), "premios.xlsx"),
        },
        content_type="multipart/form-data",
    )
    assert extension.status_code == 200
    assert "resultado del cruce anterior" in extension.get_data(as_text=True)
    extension_headers = client.post(
        "/headers",
        data={"csrf_token": token, "sheet_1": "0", "header_1": "1"},
    )
    assert extension_headers.status_code == 200
    client.post("/clean/apply", data={"csrf_token": token})
    extended_cross = client.post(
        "/run",
        data={
            "csrf_token": token,
            "left_key_0": ["CUIL", "", "", "", ""],
            "right_key_0": ["CUIL", "", "", "", ""],
            "join_type": "inner",
            "normalize": "yes",
        },
    )
    assert extended_cross.status_code == 200
    final_result = client.post(
        "/rules/skip",
        data={"csrf_token": token},
        follow_redirects=True,
    )
    assert final_result.status_code == 200
    assert "Pagado" in final_result.get_data(as_text=True)

    log_path = next(web_app._audit.directory.glob("*.jsonl"))
    audit_records = [json.loads(line) for line in log_path.read_text().splitlines()]
    assert {record["event"] for record in audit_records} >= {
        "file_loaded",
        "columns_selected",
        "process_started",
        "process_finished",
        "export",
    }
    audit_lines = log_path.read_text().splitlines()
    assert audit_records[0]["prev_hash"] == "GENESIS"
    for index in range(1, len(audit_records)):
        expected = hashlib.sha256(audit_lines[index - 1].encode("utf-8")).hexdigest()
        assert audit_records[index]["prev_hash"] == expected
    assert "20123456789" not in "\n".join(audit_lines)


def test_cruce_incompleto_permanece_en_configuracion(client):
    home = client.get("/")
    token = _csrf(home.get_data(as_text=True))
    left = _excel_bytes("Altas", [["20123456789", "Activo"]])
    right = _excel_bytes("Control", [["20123456789", "OK"]])
    client.post(
        "/upload",
        data={"csrf_token": token, "files": [(io.BytesIO(left), "altas.xlsx"), (io.BytesIO(right), "control.xlsx")]},
        content_type="multipart/form-data",
    )
    client.post("/headers", data={"csrf_token": token, "header_0": "1", "header_1": "1"})
    response = client.post(
        "/run",
        data={
            "csrf_token": token,
            "left_key_0": ["CUIL", "", "", "", ""],
            "right_key_0": ["", "", "", "", ""],
            "join_type": "inner",
        },
    )
    assert response.status_code == 200
    html = response.get_data(as_text=True)
    assert "4. Elegir columnas de cruce" in html
    assert "Completa ambos lados" in html
    assert 'option value="CUIL" selected' in html


def test_error_page_shows_reference_and_audits_not_found(client):
    response = client.get("/ruta-que-no-existe")
    assert response.status_code == 404
    html = response.get_data(as_text=True)
    assert "Referencia de error:" in html
    log_path = next(web_app._audit.directory.glob("*.jsonl"))
    records = [json.loads(line) for line in log_path.read_text().splitlines()]
    assert records[-1]["event"] == "error"
    assert records[-1]["detail"]["error_type"] == "http_404"


def test_rechaza_conexiones_no_locales(client, monkeypatch):
    blocked = client.get("/", environ_base={"REMOTE_ADDR": "192.0.2.10"})
    assert blocked.status_code == 403

    monkeypatch.setattr(
        web_app,
        "_allowed_remote_addresses",
        frozenset({ipaddress.ip_address("192.168.1.70")}),
    )
    allowed = client.get("/", environ_base={"REMOTE_ADDR": "192.168.1.70"})
    assert allowed.status_code == 200


def test_cruce_encadenado_con_tres_archivos(client):
    home = client.get("/")
    token = _csrf(home.get_data(as_text=True))
    uploads = [
        _excel_bytes("altas", [["20123456789", "Activo"]]),
        _excel_bytes("control", [["20123456789", "Validado"]]),
        _excel_bytes("premios", [["20123456789", "Pendiente"]]),
    ]
    response = client.post(
        "/upload",
        data={
            "csrf_token": token,
            "files": [
                (io.BytesIO(content), f"archivo_{index}.xlsx")
                for index, content in enumerate(uploads)
            ],
        },
        content_type="multipart/form-data",
    )
    assert response.status_code == 200
    headers = client.post(
        "/headers",
        data={
            "csrf_token": token,
            "header_0": "1",
            "header_1": "1",
            "header_2": "1",
            "sheet_0": "0",
            "sheet_1": "0",
            "sheet_2": "0",
        },
    )
    assert headers.status_code == 200
    clean = client.post(
        "/clean/apply",
        data={"csrf_token": token},
        follow_redirects=True,
    )
    assert clean.status_code == 200
    crossed = client.post(
        "/run",
        data={
            "csrf_token": token,
            "left_key_0": ["CUIL", "", "", "", ""],
            "right_key_0": ["CUIL", "", "", "", ""],
            "left_key_1": ["CUIL", "", "", "", ""],
            "right_key_1": ["CUIL", "", "", "", ""],
            "join_type": "inner",
            "normalize": "yes",
        },
    )
    assert crossed.status_code == 200
    result = client.post(
        "/rules/skip",
        data={"csrf_token": token},
        follow_redirects=True,
    )
    assert result.status_code == 200
    assert "Filas de salida" in result.get_data(as_text=True)


def test_estadisticas_configurables_usan_decimal_y_reportan_invalidos():
    frame = pd.DataFrame(
        {
            "Estado": ["A", "A", "B"],
            "Importe": ["1.234,56", "10,44", "invalido"],
        }
    )
    columns, rows = web_app._calculate_configured_statistics(
        frame,
        ["Estado"],
        [("Importe", "sum"), ("", "count_rows")],
    )
    assert columns == ["Estado", "sum: Importe", "Filas"]
    group_a = next(row for row in rows if row["Estado"] == "A")
    group_b = next(row for row in rows if row["Estado"] == "B")
    assert group_a["sum: Importe"] == "1245.00 (válidos: 2, inválidos: 0)"
    assert group_b["sum: Importe"] == "Sin valores válidos (inválidos: 1)"


def test_receta_se_ejecuta_automaticamente_y_muestra_revision(
    client, tmp_path, monkeypatch
):
    recipes_dir = tmp_path / "recipes"
    monkeypatch.setattr(web_app, "_recipes_dir", lambda: recipes_dir)
    recipe = Recipe(
        recipe_id="receta-automatica",
        name="Cruce de prueba",
        description="",
        created_at="2026-01-01T00:00:00+00:00",
        updated_at="2026-01-01T00:00:00+00:00",
        version="1.0",
        file_specs=[
            FileSpec(label="altas", expected_columns=["CUIL", "Estado"]),
            FileSpec(label="control", expected_columns=["CUIL", "Estado"]),
        ],
        cleansing_rules=[],
        join_type="inner",
        normalize_keys=True,
        cross_steps=[
            CrossStepSpec(
                left_label="altas",
                right_label="control",
                key_pairs=[("CUIL", "CUIL")],
            )
        ],
        business_rules=[],
    )
    web_app.save_recipe(recipe, recipes_dir / f"{recipe.recipe_id}.json")

    home = client.get("/")
    token = _csrf(home.get_data(as_text=True))
    result = client.post(
        "/upload",
        data={
            "csrf_token": token,
            "recipe_id": recipe.recipe_id,
            "run_mode": "automatic",
            "files": [
                (io.BytesIO(_excel_bytes("altas", [["20123456789", "Activo"]])), "altas.xlsx"),
                (io.BytesIO(_excel_bytes("control", [["20123456789", "Validado"]])), "control.xlsx"),
            ],
        },
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert result.status_code == 200
    html = result.get_data(as_text=True)
    assert "Se ejecutó automáticamente la configuración" in html
    assert "Revisar resultado antes de exportar" in html
    assert "20123456789" in html
    assert "Exportar resultado a Excel" in html
