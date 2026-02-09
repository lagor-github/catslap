"""Catslap CLI entrypoint and high-level processing logic.

This module loads JSON input and applies it to document templates to generate
final outputs. It supports single files, directories, and ZIP files containing
templates, and can optionally filter by file extension.
"""

import argparse
import json
import os
import sys
import shutil

if __package__ in (None, ""):
  repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
  if repo_root not in sys.path:
    sys.path.insert(0, repo_root)

from catslap.utils import file as file_util
from catslap.utils import zip as zip_util
from catslap.docx.document import WordDocument
from catslap.xlsx.document import ExcelDocument
from catslap.pptx.document import PowerPointDocument
from catslap.html.document import HtmlDocument


class Catslap:
  """High-level processor for applying JSON data to templates."""
  def __init__(self, json_map: dict):
    """Create a processor with parsed JSON data."""
    self.json_map = json_map
    self.update_toc = False
    self.pdf = False

  def set_word_output_config(self, update_toc: bool, pdf: bool = False):
    self.update_toc = update_toc
    self.pdf = pdf

  def process(self, template_file: str) -> bytes:
    """Return output final document bytes processed by template handler."""
    doc = None
    try:
      doc = Catslap.get_document(template_file)
      if not doc:
        raise ValueError(f"Unsupported template file: {template_file}")
      data = doc.get_bytes_with_json(self.json_map)
      if template_file.endswith('.docx') or template_file.endswith('.docm'):
        if not self.update_toc:
          return data
      return doc.update_toc(data, self.pdf)
    finally:
      if doc:
        doc.close()

  @staticmethod
  def get_document(template_path):
    """Return a document handler for the template path, or None if unsupported."""
    if template_path.endswith(('.docx', '.docm')):
      return WordDocument(template_path)
    if template_path.endswith('.xlsx'):
      return ExcelDocument(template_path)
    if template_path.endswith('.pptx'):
      return PowerPointDocument(template_path)
    if template_path.endswith('.html') or template_path.endswith('.js') or template_path.endswith('.txt') or template_path.endswith('.md'):
      return HtmlDocument(template_path)
    return None

  @staticmethod
  def _verbose_print(verbose: bool, message: str):
    """Print a message only when verbose mode is enabled."""
    if verbose:
      print(message)

  def process_file(self, template_file: str, output_file: str, exts: list, verbose: bool = False):
    """Process a single template file and generate its output."""
    done = True
    if exts:
      done = False
      for ext in exts:
        if template_file.lower().endswith(ext.lower()):
          done = True
          break
    if not done:
      Catslap._verbose_print(verbose, f"File filtered: {template_file}")
      return
    doc = None
    try:
      Catslap._verbose_print(verbose, f"  Process template: {template_file}")
      doc = Catslap.get_document(template_file)
      if not doc:
        Catslap._verbose_print(verbose, f"Unsupported template file: {template_file}")
        return
      doc.create_doc_with_json(output_file, self.json_map)
      Catslap._verbose_print(verbose, f"Generated: {output_file}")
    finally:
      if doc:
        doc.close()


  def process_directory(self, origin_dir: str, template_dir: str, output_dir: str, exts: list, verbose: bool = False):
    """Process a directory tree, preserving relative paths under output_dir."""
    Catslap._verbose_print(verbose, f"  Process directory: {template_dir}")
    for filename in os.listdir(template_dir):
      if filename == "__MACOSX" or filename.startswith("."):
        continue
      template_path = os.path.join(template_dir, filename)
      template_rest = template_dir[len(origin_dir):]
      if template_rest.startswith('/'):
        template_rest = template_rest[1:]
      output_path = os.path.join(output_dir, template_rest, filename)
      if os.path.isdir(template_path):
        if not os.path.exists(output_path):
          os.makedirs(output_path, exist_ok=False)
        self.process_directory(origin_dir, template_path, output_dir, exts, verbose)
        continue
      self.process_file(template_path, output_path, exts, verbose)


  def process_dir_or_file(self, template: str, output: str, exts: list, verbose: bool = False):
      """Process a template input that can be a file, directory, or ZIP."""
      if not os.path.exists(template):
        raise FileNotFoundError(f"Template directory or file not found: {template}")

      if not os.path.exists(output):
        raise FileNotFoundError(f"Output directory or file not found: {output}")

      if os.path.isfile(template) and template.lower().endswith('.zip'):
        if not os.path.isdir(output):
          raise ValueError("Output parameter must be a directory when template is a ZIP")
        Catslap._verbose_print(verbose, f"  Extracting ZIP: {template}")
        template_dir = zip_util.extract_all(template)
        try:
          Catslap._verbose_print(verbose, f"  Extracted to: {template_dir}")
          self.process_directory(template_dir, template_dir, output, exts, verbose)
        finally:
          # borrar directorio temporal
          shutil.rmtree(template_dir, ignore_errors=True)
          Catslap._verbose_print(verbose, f"  Temporal directory removed: {template_dir}")          
        return

      if os.path.isfile(template):
        template_file = template
        if os.path.isdir(output):
          output_file = os.path.join(output, os.path.basename(template_file))
        else:
          output_file = output
        self.process_file(template_file, output_file, exts, verbose)
        return

      template_dir = template
      output_dir = output

      if not os.path.isdir(template_dir):
        raise ValueError("Template parameter must be a file template or a directory with templates")

      if not os.path.isdir(output_dir):
        raise ValueError("Output parameter must be a directory when template is a directory")

      self.process_directory(template_dir, template_dir, output_dir, exts, verbose)


def _error(message: str):
  """Print a user-facing error message."""
  print(f"ERROR: {message}")

def main():
  """CLI entrypoint for catslap."""
  print()
  parser = argparse.ArgumentParser(
    prog="catslap",
    description="Document generator using JSON input"
  )
  parser.add_argument(
    "json_file",
    help="Input JSON file"
  )
  parser.add_argument(
    "template",
    help="Template file or templates directory"
  )
  parser.add_argument(
    "output",
    help="Output file or output directory"
  )
  parser.add_argument(
    "-v", "--verbose",
    action="store_true",
    help="Enable verbose output"
  )
  parser.add_argument(
    "-x", "--ext",
    nargs="*",
    default=None,
    help="File extensions to process ; separated (e.g. .docx;.xlsx). Defaults to all"
  )

  args = parser.parse_args()

  json_file = os.path.abspath(args.json_file)
  template_path = os.path.abspath(args.template)
  output_path = os.path.abspath(args.output)

  json_data = file_util.read_text(json_file, "UTF-8")
  json_map = json.loads(json_data)

  verbose = args.verbose
  ext_data = args.ext
  exts = None
  if ext_data:
    if isinstance(ext_data, list):
      if len(ext_data) == 1 and ';' in ext_data[0]:
        exts = [e for e in ext_data[0].split(';') if e]
      else:
        exts = ext_data
    else:
      exts = [e for e in ext_data.split(';') if e]

  Catslap._verbose_print(verbose, f"Input JSON: {json_file}")
  Catslap._verbose_print(verbose, f"Template: {template_path}")
  Catslap._verbose_print(verbose, f"Output: {output_path}")
  Catslap._verbose_print(verbose, f"Extensions filter: {exts if exts else 'all'}")

  catslap = Catslap(json_map)
  try:
    catslap.process_dir_or_file(args.template, args.output, exts, verbose)
  except FileNotFoundError as e:
    _error(str(e))
    return -1
  except ValueError as e:
    _error(str(e))
    return -2
  return 0


if __name__ == "__main__":
  ret = main()
  exit(ret)
