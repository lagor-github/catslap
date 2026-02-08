import subprocess
import traceback

def execute(argumentos) -> any:
  """
  Executes LibreOffice (soffice) with arguments.

  Args:
    argumentos: List of arguments for `soffice`.

  Returns:
    Popen object for the launched process.

  Raises:
    Exception: If execution fails or soffice is not found.
  """
  try:
  # Comando base de soffice
    comando = ["soffice"] + argumentos

    # Ejecutar el comando
    process = subprocess.Popen(
        comando,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    return process
  except subprocess.CalledProcessError as e:
    print(f"ERROR: No es posible completar la ejecución de soffice: {str(e)}")
    print(traceback.format_exc())
    raise Exception("ERROR: No es posible completar la ejecución de soffice: {str(e)}")
  except FileNotFoundError:
    print("ERROR: soffice no está instalado o no se encuentra en el PATH.")
    raise Exception("ERROR: soffice no está instalado o no se encuentra en el PATH.")
  except Exception as e:
    raise Exception("ERROR: {str(e)}")
