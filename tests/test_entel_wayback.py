import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import entel_wayback as ew  # noqa: E402

HTML_CARDS = """
<html><body>
<header><a>Personas</a> <a>Empresas</a></header>
<div class="plans">
  <div class="card"><h3>Plan 20 GB</h3><p>Minutos libres</p><p>SMS libres</p>
    <p>WhatsApp libre</p><span class="old">$19.990</span><span class="price">$14.990</span>
    <small>50% dcto por 6 meses</small></div>
  <div class="card"><h3>Plan Libre</h3><p>Gigas libres</p><p>Minutos libres</p>
    <p>Roaming en América incluido</p><span class="price">$ 29.990</span></div>
  <div class="card"><h3>Plan 500 MB</h3><p>300 minutos</p><p>100 SMS</p>
    <span class="price">$7.990</span></div>
</div>
<footer>Call center 600 360 0103</footer>
</body></html>
"""

HTML_JSON = """
<html><body><div id="root"></div>
<script id="__NEXT_DATA__" type="application/json">
{"props":{"plans":[{"name":"Plan 40 GB","price":"17990","gb":40},
                    {"name":"Plan 80 GB","price":21990,"gb":80}]}}
</script></body></html>
"""


def test_extract_cards():
    plans = ew.extract_plans(HTML_CARDS)
    by_price = {p.precio_clp: p for p in plans}
    assert set(by_price) == {14990, 29990, 7990}
    p20 = by_price[14990]
    assert p20.datos_gb == 20 and p20.precio_normal_clp == 19990
    assert p20.minutos_ilimitados and p20.sms_ilimitados
    assert "WhatsApp" in p20.redes_sociales_libres
    assert p20.descuento_pct == 50 and p20.meses_promocion == 6
    libre = by_price[29990]
    assert libre.datos_ilimitados and libre.roaming
    small = by_price[7990]
    assert abs(small.datos_gb - 500 / 1024) < 1e-9
    assert small.minutos == 300 and small.sms == 100


def test_extract_json():
    plans = ew.extract_plans(HTML_JSON)
    assert {(p.nombre, p.precio_clp) for p in plans} == {("Plan 40 GB", 17990), ("Plan 80 GB", 21990)}
    assert all(p.metodo == "json" for p in plans)


def test_end_to_end_excel(tmp_path, monkeypatch):
    from openpyxl import load_workbook

    caps = [
        {"timestamp": "20190103120000", "original": "https://www.entel.cl/planes/"},
        {"timestamp": "20190201080000", "original": "https://www.entel.cl/planes/"},
    ]
    monkeypatch.setattr(ew, "first_capture_per_month", lambda url, d, h: caps)
    pages = {"20190103120000": HTML_CARDS, "20190201080000": HTML_JSON}
    monkeypatch.setattr(ew, "fetch_snapshot", lambda ts, orig: pages[ts])

    out = tmp_path / "out.xlsx"
    ew.run(["www.entel.cl/planes/"], "2019", "2019", str(out))
    wb = load_workbook(out)
    assert wb.sheetnames == ["Planes", "Resumen_Mensual", "Capturas", "Notas"]
    assert wb["Planes"].max_row == 1 + 5
    resumen = list(wb["Resumen_Mensual"].iter_rows(min_row=2, values_only=True))
    assert [r[0] for r in resumen] == ["2019-01", "2019-02"]
    assert resumen[0][2] == 7990 and resumen[0][4] == 29990
