"""
Manejo seguro de los archivos que sube el usuario.

Streamlit vuelve a ejecutar el script completo cada vez que se toca un control.
El objeto `UploadedFile` sobrevive entre esas ejecuciones, pero **conserva la
posición del cursor**: si en la primera pasada se leyó con `.read()`, en la
segunda devuelve b"" y openpyxl falla con `BadZipFile: File is not a zip file`.

Por eso todo archivo subido se convierte en un buffer nuevo antes de usarlo.
"""

import io


def buffer(archivo_subido):
    """Copia el archivo a un buffer nuevo, siempre posicionado al inicio."""
    if archivo_subido is None:
        return None
    if hasattr(archivo_subido, "getvalue"):
        datos = archivo_subido.getvalue()
    else:
        try:
            archivo_subido.seek(0)
        except Exception:
            pass
        datos = archivo_subido.read()
    return io.BytesIO(datos)


def bytes_de(archivo_subido):
    """Contenido completo en bytes, sin depender de la posición del cursor."""
    if archivo_subido is None:
        return None
    if hasattr(archivo_subido, "getvalue"):
        return archivo_subido.getvalue()
    try:
        archivo_subido.seek(0)
    except Exception:
        pass
    return archivo_subido.read()


def nombre_de(archivo_subido, por_defecto=""):
    return getattr(archivo_subido, "name", por_defecto) or por_defecto
