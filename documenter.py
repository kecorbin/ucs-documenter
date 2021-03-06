#!/usr/bin/env python
import xlsxwriter
from ucsmsdk.ucshandle import UcsHandle
import yaml

TABS = []


class ColumnTracker(dict):
    """
    Tracks the length of columns in a worksheet for setting column width based on length of content
    """
    def __setitem__(self, key, value):
        if key in self:
            if value > self[key]:
                dict.__setitem__(self, key, value)
        else:
            dict.__setitem__(self, key, value)


def new_worksheet(workbook, name):
    """
    Verify the sheet name is compatible with xlsx writer, tracks sheet names for case based duplicates
    :param workbook: Workbook object
    :param name: str name of tab
    :return: worksheet object
    """
    global TABS
    name = name[0:31]
    if name.lower() in TABS:
        name += '-'
    TABS.append(name.lower())
    sheet = workbook.add_worksheet(name)
    return sheet


def create_tab(ucshandle, workbook, sheetname, cls, columns):
    """
    creates a sheet in workbook
    :param ucshandle: UcsHandle object
    :param workbook: Workbook object
    :param sheetname: str name of tab/sheet
    :param cls: str class of objects in this sheet
    :param columns: list of attributes/properties to which will be displayed as columns on the sheet
    :return:
    """
    # formatter for header cells
    header = workbook.add_format({"bold": True})
    # normal cell formatter
    outline = workbook.add_format()
    outline.set_border()
    tracker = ColumnTracker()
    sheet = new_worksheet(workbook, str(sheetname))
    row, col = 0, 0
    # write column headers
    for f in columns:
        sheet.write(row, col, f, header)
        tracker[col] = len(f)
        col += 1
    row, col = 1, 0
    # get data for sheet from class query
    mos = ucshandle.query_classid(cls)
    for i in mos:
        # create a list of attributes (columns)
        attributes = [str(getattr(i, s)) for s in columns]
        for val in attributes:
            tracker[col] = len(val)
            sheet.write(row, col, val, outline)
            col += 1

        row += 1
        col = 0

    # Autofit column width
    for c in tracker.keys():
        sheet.set_column(c, c, tracker[c] * 1.2)


def create_workbook(ucshandle, xls, tabs):
    """
    Creates Spreadsheet, calls createWorksheet for each tab, closes and saves spreadsheet
    :param ucshandle: UcsHandle instance
    :param xls: str filename of xlsx spreadsheet
    :param tabs: dict configuration dictionary, usually imported from yaml
    :return: None
    """
    workbook = xlsxwriter.Workbook(xls)
    for k in tabs:
        create_tab(ucshandle, workbook, k, tabs[k]['class'], tabs[k]['columns'])
    workbook.close()


with open('config.yaml', 'r') as config:
    config = yaml.safe_load(config)
hostname, username, password = config['host'], config['name'], config['passwd']
handle = UcsHandle(hostname, username, password)
print("Connecting to UCSM at {} as {} ......".format(hostname, username)),
if handle.login():
    print("Success")
    print "Generating Workbook file {}".format(config['filename'])
    create_workbook(handle, config['filename'], config['tabs'])
else:
    print("FAIL")
