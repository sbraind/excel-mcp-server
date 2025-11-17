# Guía de Pruebas - Excel MCP Server

## Archivo de Prueba
**Ruta:** `/Users/sebastianbrain/Desktop/Experimentos/Experimentos-1/excel-mcp-server/test.xlsx`

## Comandos para Probar en Claude Desktop

### 1. Leer información del archivo
```
Lee el archivo Excel en /Users/sebastianbrain/Desktop/Experimentos/Experimentos-1/excel-mcp-server/test.xlsx y muéstrame qué hojas tiene y cuántas filas/columnas tiene cada una.
```

### 2. Ver datos de una hoja
```
Muéstrame las primeras 10 filas de la hoja "Sales" del archivo test.xlsx
```

### 3. Leer una celda específica
```
¿Qué valor tiene la celda A1 en la hoja "Sales" del archivo test.xlsx?
```

### 4. Modificar una celda
```
Actualiza la celda A1 de la hoja "Sales" del archivo test.xlsx con el texto "REPORTE DE VENTAS 2024"
```

### 5. Agregar una nueva fila
```
Agrega una nueva fila al final de la hoja "Sales" con estos datos:
- Producto: "Laptop Gaming"
- Cantidad: 3
- Precio: 1500
- Total: 4500
```

### 6. Aplicar formato
```
Aplica formato de negrita y color de fondo azul a la primera fila (encabezados) de la hoja "Sales"
```

### 7. Crear una nueva hoja
```
Crea una nueva hoja llamada "Resumen" en el archivo test.xlsx
```

### 8. Buscar datos
```
Busca todas las celdas que contengan la palabra "Manager" en la hoja "Employees"
```

### 9. Crear un gráfico
```
Crea un gráfico de barras con las primeras 5 filas de la columna de ventas en la hoja "Sales"
```

### 10. Crear tabla dinámica
```
Crea una tabla dinámica que muestre la suma total de ventas por producto en la hoja "Sales"
```

## Notas
- Todas las modificaciones crearán un backup automático del archivo
- Puedes encadenar múltiples operaciones en un solo comando
- El servidor valida todos los cambios antes de aplicarlos
