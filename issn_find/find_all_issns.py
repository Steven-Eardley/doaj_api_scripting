"""Check the DOAJ search API for all ISSNs found in .xls workbooks"""

import xlrd
import re
import requests
import time

ISSN_REGEX = re.compile(r'^\d{4}-?\d{3}(\d|x)$', flags=re.IGNORECASE)
DOAJ_SEARCH = "https://doaj.org/api/v1/search/journals/"
WAIT_PERIOD = 0.1  # seconds


def issns_from_sheet(sheet):
    """Get any ISSNs from a spreadsheet by matching regex on every cell"""
    set_of_issns = set()
    duplicates = 0

    for r in range(0, sheet.nrows):
        for c in range(0, sheet.ncols):
            try:
                if sheet.cell(r, c).value != xlrd.empty_cell.value:
                    v = sheet.cell(r, c).value.strip()
                    if ISSN_REGEX.match(v):
                        if v in set_of_issns:
                            duplicates += 1
                        else:
                            set_of_issns.add(v)
            except AttributeError:
                # the ISSN could be a number, convert it to text and retry
                v = unicode(sheet.cell(r, c).value).strip()
                if ISSN_REGEX.match(v):
                    if v in set_of_issns:
                        duplicates += 1
                    else:
                        set_of_issns.add(v)

    return set_of_issns, duplicates


def is_issn_in_doaj(issn):
    """Check a single ISSN for a journal match using the DOAJ search API"""
    time.sleep(WAIT_PERIOD)  # hold your horses.
    resp = requests.get(DOAJ_SEARCH + "issn:" + issn)
    if resp.status_code == 200:
        return bool(int(resp.headers['x-total-count']))
    else:
        return None


def report_sheet(sheet):
    print "\tReading sheet {0}".format(sheet.name)
    issns, dups = issns_from_sheet(sheet)
    print "\t\tFound {0} ISSNs. {1} duplicate(s) on the sheet.".format(len(issns), dups)
    results = [is_issn_in_doaj(i) for i in list(issns)]
    filtered_results = [r for r in results if r is not None]
    print "\t\t{0} ISSNs are present in the DOAJ.".format(sum(filtered_results))
    if len(filtered_results) < len(results):
        print "\t\t\t* However, the DOAJ search failed on {0} ISSNs.".format(len(results) - len(filtered_results))


if __name__ == '__main__':

    print "Opening 2016 workbook DHET Accredited journal lists for publications made 2016.xls"
    workbook_2016 = xlrd.open_workbook('data/DHET Accredited journal lists for publications made 2016.xls',
                                       on_demand=True)

    # There is some overlap between the two files - only read the 2016 entries from this one.
    for n in workbook_2016.sheet_names()[:1]:
        if n.endswith('2016'):
            report_sheet(workbook_2016.sheet_by_name(n))

    print "\nOpening 2017 workbook DHET Accredited journal lists for publications to be made in 2017.xls"
    workbook_2017 = xlrd.open_workbook('data/DHET Accredited journal lists for publications to be made in 2017.xls',
                                       on_demand=True)

    # And only read 2017 entries from this one.
    for n in workbook_2017.sheet_names()[:1]:
        if n.endswith('2017'):
            report_sheet(workbook_2017.sheet_by_name(n))
