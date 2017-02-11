"""
Assess whether the DOAJ has both ISSNs for a journal, read from a .xls
Assumes one journal per row.
"""

import xlrd
import re
import requests
import time

ISSN_REGEX = re.compile(r'^\d{4}-?\d{3}(\d|x)$', flags=re.IGNORECASE)
DOAJ_SEARCH = "https://doaj.org/api/v1/search/journals/"
WAIT_PERIOD = 0.1  # seconds


def is_issn_in_doaj(issn):
    """Check a single ISSN for a journal match using the DOAJ search API"""
    time.sleep(WAIT_PERIOD)  # hold your horses.
    resp = requests.get(DOAJ_SEARCH + "issn:" + issn)
    if resp.status_code == 200:
        return bool(int(resp.headers['x-total-count']))
    else:
        return False


def issns_from_sheet_by_row(sheet):
    """Check whether one or both ISSNs in a row are in the DOAJ"""

    counts_in_doaj = []

    for r in range(0, sheet.nrows):

        set_of_issns = set()
        for c in range(0, sheet.ncols):
            try:
                if sheet.cell(r, c).value != xlrd.empty_cell.value:
                    v = sheet.cell(r, c).value.strip()
                    if ISSN_REGEX.match(v):
                        set_of_issns.add(v)
            except AttributeError:
                # the ISSN could be a number, convert it to text and retry
                v = unicode(sheet.cell(r, c).value).strip()
                if ISSN_REGEX.match(v):
                    set_of_issns.add(v)
        if len(set_of_issns) > 0:
            counts_in_doaj.append(sum([is_issn_in_doaj(i) for i in list(set_of_issns)]))

    return counts_in_doaj


def report_sheet(sheet):
    print "\tReading sheet {0}".format(sheet.name)
    c = issns_from_sheet_by_row(sheet)
    print "\t\tFound {0} rows containing ISSNs.".format(len(c))
    n_2 = c.count(2)
    n_1 = c.count(1)
    n_0 = c.count(0)
    print "\t\t{0} journals found by both ISSNs, {1} found by one, {2} found by neither.".format(n_2, n_1, n_0)


if __name__ == '__main__':

    print "Opening 2016 workbook DHET Accredited journal lists for publications made 2016.xls"
    workbook_2016 = xlrd.open_workbook('data/DHET Accredited journal lists for publications made 2016.xls',
                                       on_demand=True)

    # There is some overlap between the two files - only read the 2016 entries from this one.
    for n in workbook_2016.sheet_names():
        if n.endswith('2016'):
            report_sheet(workbook_2016.sheet_by_name(n))

    print "\nOpening 2017 workbook DHET Accredited journal lists for publications to be made in 2017.xls"
    workbook_2017 = xlrd.open_workbook('data/DHET Accredited journal lists for publications to be made in 2017.xls',
                                       on_demand=True)

    # And only read 2017 entries from this one.
    for n in workbook_2017.sheet_names():
        if n.endswith('2017'):
            report_sheet(workbook_2017.sheet_by_name(n))
