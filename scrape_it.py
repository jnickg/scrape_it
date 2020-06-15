import os
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
import win32com.client as win32

CONFIG_RESULT_DIR='results'

fpdf.set_global("SYSTEM_TTFONTS", join(os.path.dirname(__file__), join('res','fonts')))
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
wdFormatPDF = 17
wdDoNotSaveChanges = 0

com_app_word = None

visited_links = []

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
  fh = open(dest, "wb")
  fh.write(content.encode('utf8'))
  fh.close()

def dump_rtf_to_txt(dest='file_dest', content=None):
  global com_app_word
  text = ""
  with tempfile.NamedTemporaryFile("w+b", delete=False, suffix=".rtf") as c:
    c.write(content)
    c.close()
    com_app_word.Documents.Open(c.name)
    doc = com_app_word.ActiveDocument
    text = doc.Content.Text 
    doc.Close(wdDoNotSaveChanges)
    try:
      os.remove(c.name)
    except:
      pass
  return dump_txt_to_txt(dest=dest, content=text)

def dump_rtf_to_pdf(dest='file_dest', content=None):
  global com_app_word
  with tempfile.NamedTemporaryFile("w+b", delete=False, suffix=".rtf") as c:
    c.write(content)
    c.close()
    com_app_word.Documents.Open(c.name)
    doc = com_app_word.ActiveDocument
    doc.ReadOnlyRecommended = False
    doc.SaveAs(f"{dest}.pdf", FileFormat=wdFormatPDF)
    doc.Close(wdDoNotSaveChanges)
    try:
      os.remove(c.name)
    except:
      pass

def dump_doc_to_pdf(dest='file_dest', content=None):
  global com_app_word
  with tempfile.NamedTemporaryFile("w+b", delete=False, suffix=".doc") as c:
    c.write(content)
    c.close()
    com_app_word.Documents.Open(c.name)
    doc = com_app_word.ActiveDocument
    doc.ReadOnlyRecommended = False
    doc.SaveAs(f"{dest}.pdf", FileFormat=wdFormatPDF)
    doc.Close(wdDoNotSaveChanges)
    try:
      os.remove(c.name)
    except:
      pass

def dump_content(dest='file_dest', content=None, mime='application/octet-stream', dest_ext='txt'):
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
  elif ('application/msword' in mime):
    if (dest_ext=='pdf'):
      dump_doc_to_pdf(dest=dest, content=content)
    else:
      raise ValueError(f"Unsupported extension {dest_ext}")


def scrape_it(url='', mimes=[]):
  results = []
  split_url = urlsplit(url)
  domain = f"{split_url.scheme}://{split_url.netloc}"
  print(f"Accessing URL {url} from domain {domain}...")
  base_request = requests.get(url, allow_redirects=True, verify=False)
  print("Success. Parsing HTML...")
  soup = bs4.BeautifulSoup(base_request.content, features='html.parser')
  all_links = soup.findAll('a')
  print(f"Found {len(all_links)} links")
  for link in all_links:
    slug = link.get("href")
    link_url = f"{domain}/{slug}"
    link_response = requests.get(link_url, allow_redirects=True, verify=False)
    link_content_type = link_response.headers['Content-Type']
    is_desired_type = (any(m in link_content_type for m in mimes))
    if (is_desired_type):
      print(f"Found content {link_content_type} at {link_url}")
      yield (link_url, link_response.content, link_response.headers['Content-Type'])
  pass

def get_links_for(url='', response=None, visited=[], same_domain=True, domain=None):
  split_url = urlsplit(url)
  print(f"Getting {url}...")
  if response is None:
    response = requests.get(url, allow_redirects=True, verify=False)
  soup = bs4.BeautifulSoup(response.content, features='html.parser')
  all_links = soup.findAll('a')
  return [urljoin(domain, l.get("href")) for l in all_links if ((l.get("href") not in visited) and (not same_domain or domain is None or domain in urljoin(domain, l.get("href"))))]


def scrape_url(url='', response=None, mimes=[], destfmt='txt', baseurl=None, working_dir=None, executor=None):
  global visited_links
  if (url in visited_links):
    print(f"{url} | Skipping already-visited link")
    return 0

  split_url = urlsplit(url)

  if executor is None:
    raise ValueError("Can't do this!")

  if baseurl is None:
    baseurl = split_url.netloc
  domain = f"{split_url.scheme}://{split_url.netloc}"
  print(f"{url} | Accessing URL from domain {domain}...")

  if working_dir is None:
    working_dir = os.getcwd()

  # Step into results directory for this URL path
  results_dir = os.path.basename(os.path.normpath(split_url.path))
  working_dir = join(working_dir, results_dir)
  if not os.path.isdir(working_dir):
    os.makedirs(working_dir, exist_ok=True)

  # Get child links
  links = get_links_for(url, response, visited_links, same_domain=True, domain=domain)
  print(f"{url} | Found {len(links)} links total")
  visited_links.append(url)

  sub_scrape_futures = []
  # Scrape each link.
  link_responses = [requests.get(l, allow_redirects=True, verify=False) for l in links]

  #   If the link is for text/html, recurse
  sub_page_responses = [r for r in link_responses if 'text/html' in r.headers['Content-Type']]
  sub_page_responses = [r for r in sub_page_responses if r.url not in visited_links]
  print(f"{url} | Found {len(sub_page_responses)} new responses that are also HTML pages. Creating jobs...")
  for r in sub_page_responses:
    f = executor.submit(scrape_url, response=r, url=r.url, mimes=mimes, destfmt=destfmt, baseurl=baseurl, working_dir=working_dir, executor=executor)
    sub_scrape_futures.append(f)

  #   Otherwise, check if it's one of the mimes we care about
  content_responses = [r for r in link_responses if r not in sub_page_responses] # Filter known HTML
  content_responses = [r for r in content_responses if any(m in r.headers['Content-Type'] for m in mimes)] # Filter for MIME
  content_grabbed = 0
  print(f"{url} | Found {len(content_responses)} responses that are desired content. Creating jobs...")
  for r in content_responses:
    content_type = r.headers['Content-Type']
    print(f"{url} | Found content {content_type} at {r.url}")
    url_fname = os.path.basename(os.path.normpath(split_url.path))
    dump_content(dest=join(working_dir, url_fname), content=r.content, mime=content_type, dest_ext=args.desttype)
    content_grabbed += 1

  for f in concurrent.futures.as_completed(sub_scrape_futures):
    content_grabbed += f.result()

  return content_grabbed

def scrape_it_recursive(args):
  global com_app_word
  com_app_word = win32.Dispatch("Word.Application")
  com_app_word.Visible = False
  
  if not os.path.isdir(CONFIG_RESULT_DIR):
    os.makedirs(CONFIG_RESULT_DIR)

  global visited_links
  visited_links = []
  with concurrent.futures.ThreadPoolExecutor(max_workers=50) as executor:
    for url in args.url:
      print(f"Starting with {url}...")
      scrape_url(url=url, mimes=args.content, destfmt=args.desttype, working_dir=CONFIG_RESULT_DIR, executor=executor)


def scrape_it_main(args):
  global com_app_word
  com_app_word = win32.Dispatch("Word.Application")
  com_app_word.Visible = False
  for url in args.url:
    split_base_url = urlsplit(url)
    url_results = join(os.path.dirname(os.path.realpath(__file__)), join("result", os.path.basename(os.path.normpath(split_base_url.path))))
    if not os.path.isdir(url_results):
      os.makedirs(url_results)
    for link_url, content, mime in scrape_it(url=url, mimes=args.content):
      split_url = urlsplit(link_url)
      url_fname = os.path.basename(os.path.normpath(split_url.path))
      dump_content(dest=join(url_results, url_fname), content=content, mime=mime, dest_ext=args.desttype)
  com_app_word.Quit()

if __name__ == "__main__":
  parser = argparse.ArgumentParser(prog="scrape_id.py", description="A web scraper designed for peeling text files from a site")
  
  parser.add_argument('--url', type=str, nargs='+')
  parser.add_argument('--content', type=str, nargs='+')
  parser.add_argument('--desttype', type=str)

  args = parser.parse_args()

  scrape_it_recursive(args)