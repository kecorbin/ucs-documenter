#!/usr/bin/env python
import xlsxwriter
from ucsmsdk.ucshandle import UcsHandle
import yaml
"""
Script which will gather useful infromation and put it into a spreadsheet
"""

tabs = []

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

def new_worksheet(workbook,name):
    '''
    Verify the sheet name is compatible with xlsx writer, tracks sheet names for case based duplicates
    '''
    global tabs
    name = name[0:31]
    while True:
        if name.lower() in tabs:
            name += '-'
        else:
            tabs.append(name.lower())
            sheet = workbook.add_worksheet(name)
            break
    return sheet


def createWorkSheet(md, workbook, sheetname, cls, columns):
    # formatter for header cells
    header = workbook.add_format({"bold": True})
    # normal cell formatter
    outline = workbook.add_format()
    outline.set_border()
    tracker = ColumnTracker()
    sheet = new_worksheet(workbook, str(sheetname))
    row, col = 0,0
    # write column headers
    for f in columns:
        sheet.write(row, col, f, header)
        tracker[col] = len(f)
        col += 1
    row,col = 1,0
    # get data for sheet from class query
    mos = md.query_classid(cls)
    width_dict = dict()
    for i in mos:
        # create a list of attributes (columns)
        attributes = [str(getattr(i, s)) for s in columns]
        for i in attributes:
            tracker[col] = len(i)
            sheet.write(row, col, i, outline)
            col += 1

        row += 1
        col = 0

    # Autofit column width
    for c in tracker.keys():
        print("Setting Column {} to width {}".format(c, tracker[c]))
        sheet.set_column(c, c, tracker[c] * 1.2)

def CreateWorkBook(handle,xls,tabs):
    """
    Creates Spreadsheet, calls createWorksheet for each tab, closes and saves spreadsheet
    :param handle:
    :param xls:
    :param tabs:
    :return:
    """
    workbook = xlsxwriter.Workbook(xls)
    for k in tabs:
        createWorkSheet(handle,workbook, k,tabs[k]['class'], tabs[k]['columns'])
    workbook.close()


with open('config.yaml', 'r') as config:
    config = yaml.safe_load(config)
hostname, username, password = config['host'], config['name'], config['passwd']
handle = UcsHandle(hostname, username, password)
handle.login()
CreateWorkBook(handle, config['filename'], config['tabs'])