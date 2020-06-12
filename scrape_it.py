import os
from os.path import isfile, join
import bs4
import argparse
import requests
import urllib3
from urllib.parse import urlsplit
from tld import get_tld
import tempfile

def dump_content_pdf(dest='file_dest', content=''):
  pass

def dump_content_txt(dest='file_dest', content=''):
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
      yield (link_url, link_request.content)
  pass

def scrape_it_main(args):
  urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
  results = {}
  for url in args.url:
    rslt = scrape_it(url=url, mimes=args.content)
    results[url] = rslt

  for url in args.url:
    split_base_url = urlsplit(url)
    url_results = join(os.path.dirname(os.path.realpath(__file__)), join("result", os.path.basename(os.path.normpath(split_base_url.path))))
    if not os.path.isdir(url_results):
      os.makedirs(url_results)
    for link_url, content in scrape_it(url=url, mimes=args.content):
      split_url = urlsplit(link_url)
      url_fname = os.path.basename(os.path.normpath(split_url.path))
      fh = open(join(url_results, url_fname), "wb")
      fh.write(content)
      fh.close()
      # make file for link
  pass

if __name__ == "__main__":
  parser = argparse.ArgumentParser(prog="scrape_id.py", description="A web scraper designed for peeling text files from a site")
  
  parser.add_argument('--url', type=str, nargs='+')
  parser.add_argument('--content', type=str, nargs='+')

  args = parser.parse_args()

  scrape_it_main(args)