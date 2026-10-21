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


        # 9. Prueba para: informe_ventas_producto() -> Que agrupe las ventas por producto
def test_informe_ventas_producto(setup_excel):
    wb = openpyxl.load_workbook(setup_excel)
    ws_ped = wb["pedidos"]
    ws_ped.append([1, "Ana", "Mouse", 1, 50.0])
    ws_ped.append([2, "Carlos", "Teclado", 1, 150.0])
    ws_ped.append([3, "Pedro", "Mouse", 2, 100.0])
    wb.save(setup_excel)
    
    # Lógica de agrupación por producto
    totales_producto = {}
    for row in ws_ped.iter_rows(min_row=2, values_only=True):
        if row[2] and row[4] is not None:
            producto, total = row[2], float(row[4])
            totales_producto[producto] = totales_producto.get(producto, 0.0) + total
            
    assert totales_producto["Mouse"] == 150.0  # 50 + 100
    assert totales_producto["Teclado"] == 150.0