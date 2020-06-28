import os
import sys
import atexit
from os.path import isfile, join
import threading
import concurrent.futures
import bs4
import argparse
import requests
import urllib3
from urllib.parse import urlsplit
from urllib.parse import urljoin
from tld import get_tld
import tempfile
import fpdf
from fpdf import FPDF
import pythoncom
import win32com.client as win32
import time
from multiprocessing.pool import ThreadPool

CONFIG_RESULT_DIR='results'

rslt_key_doc = 'doc'
fpdf.set_global("SYSTEM_TTFONTS", join(os.path.dirname(__file__), join('res','fonts')))
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
wdFormatPDF = 17
wdNoProtection = -1 # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdprotectiontype?view=word-pia
wdDoNotSaveChanges = 0
wdDialogFileOpen = 80
wdAlertsNone = 0
wdOpenFormatDocument = 1

com_app_word = None
com_app_pool_sz = 1
com_app_pool = ThreadPool(processes=com_app_pool_sz)

def com_app_word_reopen():
  global com_app_word
  if com_app_word is not None:
    # print("Nabbing dialog box thing")
    # dialog = com_app_word.Dialogs.Item(wdDialogFileOpen)
    # dialog.Execute()
    try:
      print("Quitting MS Word...")
      com_app_word.Quit(0, 1, False)
    except:
      print("Failed to close")
      pass
    com_app_word = None
  print("Launching MS Word...")
  com_app_word = win32.Dispatch("Word.Application")
  com_app_word.DisplayAlerts = wdAlertsNone
  com_app_word.Visible = False

visited_links = []

def open_doc_thread(app_id, fname):
  pythoncom.CoInitialize()
  app = win32.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(app_id, pythoncom.IID_IDispatch))
  doc = app.Documents.Open(fname, PasswordDocument='wrong password', ReadOnly=True)
  doc.Activate()
  doc_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, doc)
  return doc_id

def open_doc_in_thread(app, fname, pool):
  app_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, app)
  return pool.apply_async(open_doc_thread, (app_id, fname))

def com_open_doc(fname):
  global com_app_word
  global com_app_pool
  doc = None
  try:
    print(f"Attempting to open {fname}...")
    doc_promise = open_doc_in_thread(com_app_word, fname, com_app_pool)
    doc = doc_promise.wait(2.5)
    if (doc_promise.ready() is False):
      raise ValueError(f"Failed to open {fname} (timeout).")
    doc_id = doc_promise.get() # may re-raise exception
    doc = com_app_word.ActiveDocument # We should probably marshal this from the thread, but whatever
    if (doc.HasPassword):
      raise PermissionError(f"Document {fname} is password-protected. Can copy text but CANNOT use COM to file-convert!")
    if (com_app_word.ProtectedViewWindows.Count > 0):
      print(f"Editing protected-view document: {fname}")
      doc = com_app_word.ActiveProtectedViewWindow.Edit()
    doc.ReadOnlyRecommended = False
    if (doc.ProtectionType != wdNoProtection):
      print(f"Unprotecting protected document: {fname}")
      doc.Unprotect()
  except:
    if (doc is not None):
      try:
        doc.Close(wdDoNotSaveChanges)
      except:
        print("Failed to close a document!")
      doc = None
    com_app_word_reopen()
    raise
  return doc

def requests_get_with_wait(url):
  time.sleep(0.5)
  print(f"Getting url: {url}...")
  return requests.get(url, allow_redirects=True, verify=False)

def dump_txt_to_pdf(dest='file_dest', content=''):
  dump_txt_to_txt(f"{dest}.txt", content)
  pdf = FPDF()
  pdf.add_page()
  pdf.add_font("NotoSans", style="", fname="NotoSans-Regular.ttf", uni=True)
  pdf.add_font("NotoSans", style="B", fname="NotoSans-Bold.ttf", uni=True)
  pdf.add_font("NotoSans", style="I", fname="NotoSans-Italic.ttf", uni=True)
  pdf.add_font("NotoSans", style="BI", fname="NotoSans-BoldItalic.ttf", uni=True)

  content = content.replace("\t", "")
  pdf.set_font("NotoSans", size=10)
  pdf.write(10, content.encode('utf8').decode('utf8'))
  pdf.output(f"{dest}.pdf")

def dump_txt_to_txt(dest='file_dest', content=''):
  fh = open(f"{dest}.txt", "wb")
  fh.write(content)
  fh.close()

def dump_rtf_to_txt(dest='file_dest', content=None):
  text = ""
  with tempfile.NamedTemporaryFile("w+b", delete=False, suffix=".rtf") as c:
    c.write(content)
    c.close()
    doc = com_open_doc(c.name)
    text = doc.Content.Text 
    doc.Close(wdDoNotSaveChanges)
    try:
      os.remove(c.name)
    except:
      pass
  return dump_txt_to_txt(dest=dest, content=text.encode('utf8'))

def dump_rtf_to_pdf(dest='file_dest', content=None):
  with tempfile.NamedTemporaryFile("w+b", delete=False, suffix=".rtf") as c:
    c.write(content)
    c.close()
    doc = com_open_doc(c.name)
    doc.SaveCopyAs(f"{dest}.pdf", FileFormat=wdFormatPDF)
    doc.Close(wdDoNotSaveChanges)
    try:
      os.remove(c.name)
    except:
      pass

def dump_doc_to_pdf(dest='file_dest', content=None):
  with tempfile.NamedTemporaryFile("w+b", delete=False, suffix=".doc") as c:
    c.write(content)
    c.close()
    doc = com_open_doc(c.name)
    doc.SaveCopyAs(f"{dest}.pdf", FileFormat=wdFormatPDF)
    doc.Close(wdDoNotSaveChanges)
    try:
      os.remove(c.name)
    except:
      pass

def dump_doc_to_txt(dest='file_dest', content=None):
  text = ""
  with tempfile.NamedTemporaryFile("w+b", delete=False, suffix=".doc") as c:
    c.write(content)
    c.close()
    doc = com_open_doc(c.name)
    text = doc.Content.Text 
    doc.Close(wdDoNotSaveChanges)
    try:
      os.remove(c.name)
    except:
      pass
  return dump_txt_to_txt(dest=dest, content=text.encode('utf8'))

def dump_content_internal(dest='file_dest', content=None, mime='application/octet-stream', dest_ext='txt'):
  if ('text/plain' in mime):
    if (dest_ext=='pdf'):
      return dump_txt_to_pdf(dest=dest, content=content)
    elif(dest_ext=='txt'):
      return dump_txt_to_txt(dest=dest, content=content)
    else:
      raise ValueError(f"Unsupported extension {dest_ext}")
  elif ('application/rtf' in mime or 'text/rtf' in mime):
    if (dest_ext=='pdf'):
      return dump_rtf_to_pdf(dest=dest, content=content)
    elif(dest_ext=='txt'):
      return dump_rtf_to_txt(dest=dest, content=content)
    else:
      raise ValueError(f"Unsupported extension {dest_ext}")
  elif ('application/msword' in mime or 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' in mime):
    if (dest_ext=='pdf'):
      dump_doc_to_pdf(dest=dest, content=content)
    elif(dest_ext=='txt'):
      return dump_doc_to_txt(dest=dest, content=content)
    else:
      raise ValueError(f"Unsupported extension {dest_ext}")

def dump_content(dest='file_dest', content=None, mime='application/octet-stream', dest_ext='txt'):
  dumped = 0
  try:
    dump_content_internal(dest=dest, content=content, mime=mime, dest_ext=dest_ext)
    dumped += 1
  except:
    print(f"Failed to dump content to file {dest}.{dest_ext} due to error: {sys.exc_info()}")
  return dumped

def dump_file_exists(dest='file_dest', dest_ext='txt'):
  return os.path.isfile(f"{dest}.{dest_ext}")

def get_links_for(url='', response=None, visited=[], same_domain=True, domain=None):
  split_url = urlsplit(url)
  print(f"Getting {url}...")
  if response is None:
    response = requests_get_with_wait(url)
  soup = bs4.BeautifulSoup(response.content, features='html.parser')
  all_links = soup.findAll('a')
  return [urljoin(domain, l.get("href")) for l in all_links if ((l.get("href") not in visited) and (not same_domain or domain is None or domain in urljoin(domain, l.get("href"))))]

def scrape_url_recursive(url='', mimes=[], destfmt='txt', domain=None, working_dir=None, executor=None):
  global visited_links
  # Normalize link via urljoin
  normalized_url = urljoin(domain, url)
  split_url = urlsplit(normalized_url)

  # Check if normalized link has been visited. If so, return
  if (normalized_url in visited_links):
    print(f"{url} | Skipping already-visited URL!")
    return 0

  # If not, create directory and update working dir
  response_dest = os.path.basename(os.path.normpath(split_url.path)) # this may be a directory or filename

  # Check if a dump for this link is already present
  if (dump_file_exists(dest=join(working_dir, response_dest), dest_ext=destfmt)):
    print(f"{url} | Skipping already-visited URL (file downloaded in previous run)!")
    return 0

  # Get URL and add it to visited links
  r = requests_get_with_wait(normalized_url)
  visited_links.append(normalized_url)

  # Check if it is mime content. If so, save it to working dir, and return
  r_headers_content_type = r.headers['Content-Type']
  if (any(m in r_headers_content_type for m in mimes)):
    return dump_content(dest=join(working_dir, response_dest), content=r.content, mime=r_headers_content_type, dest_ext=destfmt)
  # If response is text/html, recurse through links
  elif ('text/html' in r_headers_content_type):
    working_dir = join(working_dir, response_dest)
    if not os.path.isdir(working_dir):
      os.makedirs(working_dir, exist_ok=True)
    links = get_links_for(url, r, visited_links, same_domain=True, domain=domain)
    content_grabbed = 0
    for l in links:
      content_grabbed += scrape_url_recursive(url=l, mimes=mimes, destfmt=destfmt, domain=domain, working_dir=working_dir, executor=executor)
    return content_grabbed
  else:
    print(f"Skipping content of type {r_headers_content_type}")
    return 0

def scrape_it_atexit():
  global com_app_word
  if (com_app_word is not None):
    try:
      com_app_word.Quit(0, 1, False)
    except:
      pass
    com_app_word = None
  pythoncom.CoUninitialize()

def scrape_it_recursive(args):
  global com_app_word
  pythoncom.CoInitialize()
  com_app_word_reopen()
  print(f"Using MS word version: {com_app_word.Version}")
  atexit.register(scrape_it_atexit)
  
  if not os.path.isdir(CONFIG_RESULT_DIR):
    os.makedirs(CONFIG_RESULT_DIR, exist_ok=True)

  global visited_links
  visited_links = []
  with concurrent.futures.ThreadPoolExecutor(max_workers=50) as executor:
    for url in args.url:
      print(f"Starting with {url}...")
      split_url = urlsplit(url)
      domain = f"{split_url.scheme}://{split_url.netloc}"
      scrape_url_recursive(url=url, mimes=args.content, destfmt=args.desttype, domain=domain, working_dir=CONFIG_RESULT_DIR, executor=executor)


if __name__ == "__main__":
  parser = argparse.ArgumentParser(prog="scrape_id.py", description="A web scraper designed for peeling text files from a site")
  
  parser.add_argument('--url', type=str, nargs='+')
  parser.add_argument('--content', type=str, nargs='+')
  parser.add_argument('--desttype', type=str)

  args = parser.parse_args()

  scrape_it_recursive(args)