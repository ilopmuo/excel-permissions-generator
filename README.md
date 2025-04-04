# Generador de Excel de Permisos

Una aplicación de escritorio que permite generar archivos Excel con permisos de objetos, utilizando checkboxes y ordenamiento alfabético.

## Características

- Interfaz gráfica de usuario intuitiva
- Generación de Excel con formato profesional
- Checkboxes para representar permisos
- Ordenamiento alfabético automático
- Formato de columnas optimizado
- Ejemplo incluido para facilitar el uso

## Requisitos para Desarrollo

- Python 3.7 o superior
- pandas
- openpyxl
- tkinter (incluido en Python)

## Instalación para Desarrollo

1. Clonar el repositorio:
```bash
git clone [URL_DEL_REPOSITORIO]
```

2. Instalar dependencias:
```bash
pip install -r requirements.txt
```

## Uso

### Usando el Ejecutable
1. Descargar el archivo `permission_excel_gui.exe` de la sección de releases
2. Ejecutar el programa haciendo doble clic
3. Ingresar los datos en el formato especificado
4. Hacer clic en "Generar Excel"

### Para Desarrollo
1. Ejecutar el script principal:
```bash
python permission_excel_gui.py
```

## Formato de Entrada

El texto debe seguir este formato:
```
Nombre del Objeto
Checked/Not Checked    Checked/Not Checked    ...    Checked/Not Checked
```

Los permisos representan:
- Create
- Read
- Edit
- Delete
- View All Records
- Modify All Records
- View All Fields

## Generación del Ejecutable

Para generar el ejecutable:
```bash
python -m PyInstaller --onefile --windowed permission_excel_gui.py
```

## Contribuir

Las contribuciones son bienvenidas. Por favor, abre un issue para discutir los cambios propuestos.

## Licencia

[MIT License](LICENSE) 