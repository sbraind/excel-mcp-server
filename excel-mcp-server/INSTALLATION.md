# Guía de Instalación para Claude Desktop

## Paso 1: Asegúrate de que el proyecto esté compilado

```bash
cd /home/user/Experimentos/excel-mcp-server
npm install
npm run build
```

## Paso 2: Verifica la ruta del servidor

La ruta completa del servidor es:
```
/home/user/Experimentos/excel-mcp-server/dist/index.js
```

## Paso 3: Configura Claude Desktop

### En macOS:

1. Abre el archivo de configuración:
   ```bash
   nano ~/Library/Application\ Support/Claude/claude_desktop_config.json
   ```

2. Agrega esta configuración (ajusta la ruta si es necesario):
   ```json
   {
     "mcpServers": {
       "excel": {
         "command": "node",
         "args": [
           "/home/user/Experimentos/excel-mcp-server/dist/index.js"
         ]
       }
     }
   }
   ```

### En Windows:

1. Abre el archivo de configuración ubicado en:
   ```
   %APPDATA%\Claude\claude_desktop_config.json
   ```

2. Agrega esta configuración (ajusta la ruta según tu instalación):
   ```json
   {
     "mcpServers": {
       "excel": {
         "command": "node",
         "args": [
           "C:\\Users\\TuUsuario\\Experimentos\\excel-mcp-server\\dist\\index.js"
         ]
       }
     }
   }
   ```

### En Linux:

1. Abre el archivo de configuración:
   ```bash
   nano ~/.config/Claude/claude_desktop_config.json
   ```

2. Agrega esta configuración:
   ```json
   {
     "mcpServers": {
       "excel": {
         "command": "node",
         "args": [
           "/home/user/Experimentos/excel-mcp-server/dist/index.js"
         ]
       }
     }
   }
   ```

## Paso 4: Reinicia Claude Desktop

Cierra completamente Claude Desktop y vuelve a abrirlo.

## Paso 5: Verifica la instalación

1. Abre Claude Desktop
2. En el chat, deberías ver un pequeño ícono de herramientas o una indicación de que MCP servers están activos
3. Prueba con un comando como:
   ```
   List the sheets in the file /home/user/Experimentos/excel-mcp-server/test.xlsx
   ```

## Solución de Problemas

### El servidor no aparece en Claude Desktop

1. **Verifica que Node.js esté instalado:**
   ```bash
   node --version
   ```
   Debe ser versión 18 o superior.

2. **Prueba el servidor manualmente:**
   ```bash
   node /home/user/Experimentos/excel-mcp-server/dist/index.js
   ```
   Deberías ver el mensaje: `Excel MCP Server running on stdio`

3. **Verifica la sintaxis del JSON:**
   - Asegúrate de que no haya comas finales
   - Verifica que todas las llaves y corchetes estén balanceados
   - Usa un validador JSON si es necesario

4. **Revisa los logs de Claude Desktop:**
   - **macOS**: `~/Library/Logs/Claude/`
   - **Windows**: `%APPDATA%\Claude\logs\`
   - **Linux**: `~/.config/Claude/logs/`

### Error: "command not found: node"

Claude Desktop podría no encontrar Node.js. En este caso, usa la ruta completa:

**macOS/Linux:**
```bash
which node  # Para encontrar la ruta de node
```

Luego actualiza la configuración:
```json
{
  "mcpServers": {
    "excel": {
      "command": "/usr/local/bin/node",
      "args": [
        "/home/user/Experimentos/excel-mcp-server/dist/index.js"
      ]
    }
  }
}
```

**Windows:**
```json
{
  "mcpServers": {
    "excel": {
      "command": "C:\\Program Files\\nodejs\\node.exe",
      "args": [
        "C:\\Users\\TuUsuario\\Experimentos\\excel-mcp-server\\dist\\index.js"
      ]
    }
  }
}
```

### Permisos en Linux/macOS

Si tienes problemas de permisos:
```bash
chmod +x /home/user/Experimentos/excel-mcp-server/dist/index.js
```

## Configuración Avanzada

### Múltiples servidores MCP

Si ya tienes otros servidores MCP configurados:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": [
        "/home/user/Experimentos/excel-mcp-server/dist/index.js"
      ]
    },
    "otro-servidor": {
      "command": "...",
      "args": ["..."]
    }
  }
}
```

## Verificación Final

Una vez configurado correctamente, Claude podrá:
- ✅ Leer archivos Excel
- ✅ Crear y modificar hojas de cálculo
- ✅ Aplicar formatos y estilos
- ✅ Realizar búsquedas y filtrados
- ✅ Manipular datos con fórmulas

Prueba con el archivo de ejemplo incluido:
```
Can you read the Sales sheet from /home/user/Experimentos/excel-mcp-server/test.xlsx?
```
