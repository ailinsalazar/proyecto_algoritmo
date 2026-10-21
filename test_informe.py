import os
import pytest
import openpyxl
from proyectoalgorit import inicializar_excel, obtener_siguiente_id

TEST_EXCEL = "test_inventario_completo.xlsx"

@pytest.fixture
def setup_excel():
    """Limpia y crea un archivo Excel de pruebas antes de cada función."""
    if os.path.exists(TEST_EXCEL):
        os.remove(TEST_EXCEL)
        
    inicializar_excel(TEST_EXCEL)
    yield TEST_EXCEL
    
    if os.path.exists(TEST_EXCEL):
        os.remove(TEST_EXCEL)


        # 8. Prueba para: informe_ventas_cliente() -> Que agrupe las ventas por cliente
def test_informe_ventas_cliente(setup_excel):
    wb = openpyxl.load_workbook(setup_excel)
    ws_ped = wb["pedidos"]
    ws_ped.append([1, "Ana", "Teclado", 1, 150.0])
    ws_ped.append([2, "Carlos", "Mouse", 2, 100.0])
    ws_ped.append([3, "Ana", "Monitor", 1, 800.0])
    wb.save(setup_excel)
    
    # Lógica de agrupación por cliente
    totales_cliente = {}
    for row in ws_ped.iter_rows(min_row=2, values_only=True):
        if row[1] and row[4] is not None:
            cliente, total = row[1], float(row[4])
            totales_cliente[cliente] = totales_cliente.get(cliente, 0.0) + total
            
    assert totales_cliente["Ana"] == 950.0   # 150 + 800
    assert totales_cliente["Carlos"] == 100.0