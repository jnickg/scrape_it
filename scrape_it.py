import os
from os.path import isfile, join
import bs4
import argparse
import requests
import urllib3
from urllib.parse import urlsplit
from tld import get_tld
import tempfile
import fpdf
from fpdf import FPDF

fpdf.set_global("SYSTEM_TTFONTS", join(os.path.dirname(__file__), join('res','fonts')))

def dump_content_pdf(dest='file_dest', content=''):
  dump_content_txt(f"{dest}.txt", content)
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

def dump_content_txt(dest='file_dest', content=''):
  fh = open(dest, "wb")
  fh.write(content.encode('utf8'))
  fh.close()
  pass

def dump_content(dest='file_dest', content='', dest_ext='txt'):
  if (dest_ext=='pdf'):
    return dump_content_pdf(dest=dest, content=content)
  elif(dest_ext=='txt'):
    return dump_content_txt(dest=dest, content=content)
  else:
    raise ValueError(f"Unknown extension {dest_ext}")

def scrape_it(url='', mimes=[]):
  results = []
  split_url = urlsplit(url)
  domain = f"{split_url.scheme}://{split_url.netloc}"
  base_request = requests.get(url, allow_redirects=True, verify=False)
  soup = bs4.BeautifulSoup(base_request.content)
  for link in soup.findAll("a"):
    slug = link.get("href")
    link_url = f"{domain}/{slug}"
    link_request = requests.get(link_url, allow_redirects=True, verify=False)
    link_content_type = link_request.headers['Content-Type']
    is_desired_type = (any(m in link_content_type for m in mimes))
    if (is_desired_type):
      print(f"Found content {link_content_type} at {link_url}")
      yield (link_url, link_request.text)
  pass

def scrape_it_main(args):
  urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
  for url in args.url:
    split_base_url = urlsplit(url)
    url_results = join(os.path.dirname(os.path.realpath(__file__)), join("result", os.path.basename(os.path.normpath(split_base_url.path))))
    if not os.path.isdir(url_results):
      os.makedirs(url_results)
    for link_url, content in scrape_it(url=url, mimes=args.content):
      split_url = urlsplit(link_url)
      url_fname = os.path.basename(os.path.normpath(split_url.path))
      dump_content(dest=join(url_results, url_fname), content=content, dest_ext=args.desttype)
  pass

if __name__ == "__main__":
  parser = argparse.ArgumentParser(prog="scrape_id.py", description="A web scraper designed for peeling text files from a site")
  
  parser.add_argument('--url', type=str, nargs='+')
  parser.add_argument('--content', type=str, nargs='+')
  parser.add_argument('--desttype', type=str)

  args = parser.parse_args()

  scrape_it_main(args)