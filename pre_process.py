# -*- coding: utf-8 -*-
# author:Haochun Wang

import xlrd, xlwt, time


def sale_to_trans(s_addr, d_addr):
    '''
    This function is to generate transactions for each purchase for one customer from all sales records.
    Write a new csv file including records like this: customer_id -- product1 -- product2--.....
    :param s_addr: The source address of the sale.xlsx file
    :param d_addr: The destination address of output file
    :return: no return
    '''
    start = time.clock()
    transacionlst = []
    customerid = 0
    customercursor = 1
    newbook = xlwt.Workbook(encoding='utf-8',style_compression=0)
    newsheet = newbook.add_sheet('sheet1', cell_overwrite_ok=True)
    items = xlrd.open_workbook(s_addr).sheet_by_index(0)
    num_of_rows = items.nrows
    currentrow = 1
    newsheet.write(0, 0, 'keyC')
    newsheet.write(0, 1, 'transaction(s)')
    for i in xrange(1, num_of_rows):
        if items.row_values(i)[1] == customerid:
            newsheet.write(currentrow, customercursor, items.row_values(i)[2])
            customercursor += 1
        elif items.row_values(i)[1] != customerid and customerid != 0:
            customerid = items.row_values(i)[1]
            currentrow += 1
            customercursor = 1
            newsheet.write(currentrow, 0, customerid)
            newsheet.write(currentrow, customercursor, items.row_values(i)[2])
            customercursor += 1
        elif items.row_values(i)[1] != customerid and customerid == 0:
            customerid = items.row_values(i)[1]
            newsheet.write(currentrow, 0, customerid)
            newsheet.write(currentrow, customercursor, items.row_values(i)[2])
            customercursor += 1
    newbook.save(d_addr)
    '''
    print items.row_values(1)[1]
    print items.row_values(1)[2]
    print type(items.row_values(1)[1])
    print type(items.row_values(1)[2])
    '''

    end = time.clock()
    print "Done! With a time comsuming of %0.2f seconds" % (end - start)
    return
sale_to_trans("/Users/hcwang/Desktop/dm/sales.xlsx", "/Users/hcwang/Desktop/dm/sales_mod.xls")


