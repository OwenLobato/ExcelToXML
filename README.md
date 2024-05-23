# Excel to XML Converter

Conversor de archivo de Excel a un archivo XML.

## Requisitos previos

- Node.js (versión 18.19.1 o superior)
- npm (administrador de paquetes de Node.js)

## Instalación

1. Clona el repo
2. Navega al directorio
3. Instala las dependencias ejecutando:

```bash
  npm install
```

## Uso

1. Asegúrate de tener un archivo de Excel en la raíz y especifica el nombre en `inputFile`. Este archivo debe contener una tabla con los datos que deseas convertir a XML.

2. Especifica el nombre de la hoja en la que se encuentran los datos en `sheetName`.

3. Especifica las columnas que quieras de la tabla de la tabla en `columnHeaders`.

4. En la iteracion de las filas, especifica cómo quieres que se nombren los atributos,

```javascript
  customObject.ele('object-attribute', { 'attribute-id': 'columnaA' }, getCellValue(sheet, columnHeaders.COL_1, i));
```

5. Para ejecutar el script y generar el archivo XML, utiliza el siguiente comando:

   ```bash
   npm run start
   ```

   Esto leerá los datos del archivo de Excel en `inputFile` y generará un archivo XML llamado como se declaro en `outputFile` en el mismo directorio.
