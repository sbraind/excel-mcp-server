# 🤝 Modo Colaborativo - Excel MCP Server

## ✨ Nueva Funcionalidad: Trabajo en Tiempo Real

El Excel MCP Server ahora soporta **trabajo colaborativo en tiempo real** entre Claude Desktop y Microsoft Excel.

## 🎯 Cómo Funciona

### Modo Colaborativo (Excel Abierto)
Cuando tienes un archivo Excel **ABIERTO** en Microsoft Excel:
- ✅ Claude detecta automáticamente que el archivo está abierto
- ✅ Usa **AppleScript** para modificar Excel directamente
- ✅ Los cambios aparecen **INSTANTÁNEAMENTE** en tu pantalla
- ✅ Puedes ver y trabajar mientras Claude hace cambios
- ✅ **Verdadero trabajo en equipo**

### Modo File-Based (Excel Cerrado)
Cuando el archivo Excel está **CERRADO**:
- ✅ Claude usa ExcelJS para modificar el archivo
- ✅ Los cambios se guardan en disco
- ✅ Debes abrir Excel para ver los cambios

## 🚀 Cómo Probar

### Paso 1: Preparar el archivo
```bash
# El archivo de prueba ya existe
/Users/sebastianbrain/Desktop/Experimentos/Experimentos-1/excel-mcp-server/test.xlsx
```

### Paso 2: Abrir el archivo en Excel
1. Abre **Microsoft Excel**
2. Abre el archivo `test.xlsx`
3. **Deja Excel abierto** con el archivo visible

### Paso 3: Probar desde Claude Desktop

#### Test 1: Actualizar una celda
```
Actualiza la celda A1 de la hoja "Sales" del archivo /Users/sebastianbrain/Desktop/Experimentos/Experimentos-1/excel-mcp-server/test.xlsx con el texto "MODO COLABORATIVO ACTIVO"
```

**Resultado esperado:**
- ✅ Verás el texto aparecer INMEDIATAMENTE en Excel
- ✅ Claude responderá con: `method: 'applescript'`
- ✅ Mensaje: "Changes are visible immediately in Excel"

#### Test 2: Agregar una fila
```
Agrega una nueva fila al final de la hoja "Sales" con estos datos:
- Producto: "MacBook Pro"
- Cantidad: 2
- Precio: 2500
- Total: 5000
```

**Resultado esperado:**
- ✅ La nueva fila aparece INSTANTÁNEAMENTE
- ✅ Puedes scrollear y verla inmediatamente
- ✅ Claude responde con: `method: 'applescript'`

#### Test 3: Múltiples cambios consecutivos
```
Actualiza estas celdas en la hoja "Sales":
- A2: "Producto Actualizado"
- B2: 100
- C2: 50
```

**Resultado esperado:**
- ✅ Cada cambio aparece en tiempo real
- ✅ Ves las celdas actualizándose una por una
- ✅ Sin necesidad de cerrar/abrir Excel

### Paso 4: Comparar con modo cerrado

1. **Cierra Excel** (Cmd+Q)
2. Ejecuta el mismo comando:
```
Actualiza la celda A1 de la hoja "Sales" del archivo test.xlsx con el texto "MODO FILE-BASED"
```

**Resultado esperado:**
- ✅ Claude responde con: `method: 'exceljs'`
- ✅ Mensaje: "File updated. Open in Excel to see changes."
- ✅ Abre Excel → Verás el cambio

## 🎬 Escenarios de Uso

### Escenario 1: Análisis colaborativo
- Tú: Miras el Excel abierto
- Claude: Actualiza fórmulas y datos
- Resultado: Ves los cálculos actualizándose en vivo

### Escenario 2: Data entry asistido
- Tú: Identificas qué datos faltan
- Claude: Llena las celdas mientras observas
- Resultado: Validación inmediata

### Escenario 3: Corrección en tiempo real
- Tú: "Claude, ese valor en B5 está mal"
- Claude: Lo corrige instantáneamente
- Resultado: Feedback loop rápido

## ⚡ Ventajas del Modo Colaborativo

1. **Feedback instantáneo**: Ves los cambios mientras suceden
2. **Sin conflicts**: Excel maneja el archivo, Claude usa su API
3. **Trabajo fluido**: No necesitas cerrar/abrir
4. **Validación inmediata**: Verificas cambios al instante
5. **Productividad++**: Flujo de trabajo continuo

## 🔧 Detalles Técnicos

### Detección Automática
El servidor detecta automáticamente:
```
1. ¿Excel está corriendo? → Si no, usa ExcelJS
2. ¿El archivo está abierto? → Si no, usa ExcelJS
3. Todo OK → Usa AppleScript
```

### AppleScript vs ExcelJS

| Característica | AppleScript | ExcelJS |
|----------------|-------------|---------|
| Velocidad visible | Instantánea | Al abrir |
| Requiere Excel | Sí | No |
| Trabajo colaborativo | ✅ | ❌ |
| Funciona offline | Solo si Excel abierto | ✅ |

### Operaciones Soportadas (v1)

Actualmente soportan modo colaborativo:
- ✅ `excel_update_cell` - Actualizar celda
- ✅ `excel_add_row` - Agregar fila

Próximamente:
- 🔜 `excel_write_range` - Escribir rango
- 🔜 `excel_create_sheet` - Crear hoja
- 🔜 `excel_delete_sheet` - Eliminar hoja
- 🔜 Más operaciones...

## 🐛 Troubleshooting

### "Method: exceljs" cuando Excel está abierto
**Problema**: El archivo puede tener un nombre diferente
**Solución**: Verifica que el archivo se llame exactamente `test.xlsx`

### Cambios no aparecen
**Problema**: AppleScript puede tener delay
**Solución**: Espera 1-2 segundos, debería aparecer

### Error de AppleScript
**Problema**: Permisos de macOS
**Solución**: Ve a System Settings → Privacy & Security → Automation → Allow Claude Desktop to control Excel

## 📝 Notas

- El modo colaborativo funciona SOLO en macOS (usa AppleScript)
- Requiere Microsoft Excel instalado
- Claude Desktop debe tener permisos para controlar Excel
- Los cambios se guardan automáticamente en Excel

## 🎉 Disfruta el Trabajo en Equipo!

Ahora puedes trabajar verdaderamente en conjunto con Claude en tus hojas de cálculo. ¡Es como tener un asistente que trabaja contigo en tiempo real!
