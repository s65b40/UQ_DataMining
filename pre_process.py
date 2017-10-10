# -*- coding: utf-8 -*-
# author:Haochun Wang

import xlrd, xlwt, time


class PreProcess:
    def __init__(self):
        pass

    def sale_to_trans(self, s_addr, d_addr):
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

    def sale_to_trans_time(self, s_addr, d_addr):
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
        customercursor = 2
        newbook = xlwt.Workbook(encoding='utf-8',style_compression=0)
        newsheet = newbook.add_sheet('sheet1', cell_overwrite_ok=True)
        items = xlrd.open_workbook(s_addr).sheet_by_index(0)
        num_of_rows = items.nrows
        currentrow = 1
        newsheet.write(0, 0, 'keyC')
        newsheet.write(0, 1, 'keyT')
        newsheet.write(0, 2, 'transaction(s)')
        for i in xrange(1, num_of_rows):
            if items.row_values(i)[1] == customerid:
                # the customer is the same as last one
                newsheet.write(currentrow, customercursor, items.row_values(i)[2])
                customercursor += 1
            elif items.row_values(i)[1] != customerid and customerid != 0:
                # the customer is not same as last one and not the first row
                customerid = items.row_values(i)[1]
                currentrow += 1
                customercursor = 2
                newsheet.write(currentrow, 0, customerid)
                newsheet.write(currentrow, 1, items.row_values(i)[0])
                newsheet.write(currentrow, customercursor, items.row_values(i)[2])
                customercursor += 1
            elif items.row_values(i)[1] != customerid and customerid == 0:
                # the customer is not same as last one and is the first row
                customerid = items.row_values(i)[1]
                newsheet.write(currentrow, 0, customerid)
                newsheet.write(currentrow, 1, items.row_values(i)[0])
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
    '''
    def sales_to_pieces(s_addr, d_addr):
        start = time.clock()
        newbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
        newsheet = newbook.add_sheet('sheet1', cell_overwrite_ok=True)
        items = xlrd.open_workbook(s_addr).sheet_by_index(0)
    
        for i in xrange(1000):
    
    
    
        newbook.save(d_addr)
    '''

    def mapping(self, s_addr, tr_addr):
        '''
        Because the products in the transitions are not classified by the type of product
        :param s_addr:
        :param d_addr:
        :return:
        '''
        map_dic = {}        #{type:[list:keys of products}
        type_dic = {}       #{type:num_order}
        pro_dic = {}        #{product:type num}
        #newbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
        #newsheet = newbook.add_sheet('sheet1', cell_overwrite_ok=True)
        items = xlrd.open_workbook(s_addr).sheet_by_index(0)
        for i in xrange(1, items.nrows):
            key_p = int(items.row_values(i)[0])
            #print key_p
            type_p = items.row_values(i)[8]
            #print type_p
            if map_dic.has_key(type_p):
                map_dic[type_p].append(key_p)
                #print type(map_dic[type_p])
            else:
                map_dic[type_p] = [key_p]
        #print map_dic
        #print len(map_dic)
        i = 0
        for j in map_dic:
            type_dic[j] = i
            for k in map_dic[j]:
                pro_dic[k] = i
            i += 1
            #print map_dic[j]
        print type_dic
        print pro_dic


        trans = xlrd.open_workbook(tr_addr).sheet_by_index(0)
        newbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
        newsheet = newbook.add_sheet('sheet1', cell_overwrite_ok=True)
        newsheet2 = newbook.add_sheet('sheet2', cell_overwrite_ok=True)
        newsheet3 = newbook.add_sheet('sheet3', cell_overwrite_ok=True)
        newsheet4 = newbook.add_sheet('sheet4', cell_overwrite_ok=True)
        #wrow = 0
        for i in range(1, trans.nrows):
            j = 1
            try:
                #print trans.row_values(i)[j]
                while isinstance(trans.row_values(i)[j], float):
                    #print pro_dic[int(trans.row_values(i)[j])]
                    newsheet.write(i, j, pro_dic[int(trans.row_values(i)[j])])
                    j += 1
            except:
                pass
        '''
        map_dic = {}        #{type:[list:keys of products}
        type_dic = {}       #{type:num_order}
        pro_dic = {}        #{product:type num}
        '''
        a = 0
        for i in map_dic:
            b = 0
            newsheet2.write(a, b, i)
            b += 1
            for j in map_dic[i]:
                newsheet2.write(a, b, j)
                b += 1
            a += 1
        a = 0
        for i in type_dic:
            newsheet3.write(a, 0, i)
            newsheet3.write(a, 1, type_dic[i])
            a += 1

        a = 0
        for i in pro_dic:
            newsheet4.write(a, 0, i)
            newsheet4.write(a, 1, pro_dic[i])
            a += 1
        newbook.save("/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/processed_data/sales_m_mod.xls")
        return 0

    def mapping_t(self, s_addr, tr_addr):
        '''
        Because the products in the transitions are not classified by the type of product
        :param s_addr:
        :param d_addr:
        :return:
        '''
        map_dic = {}        #{type:[list:keys of products}
        type_dic = {}       #{type:num_order}
        pro_dic = {}        #{product:type num}
        #newbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
        #newsheet = newbook.add_sheet('sheet1', cell_overwrite_ok=True)
        items = xlrd.open_workbook(s_addr).sheet_by_index(0)
        for i in xrange(1, items.nrows):
            key_p = int(items.row_values(i)[0])
            #print key_p
            type_p = items.row_values(i)[8]
            #print type_p
            if map_dic.has_key(type_p):
                map_dic[type_p].append(key_p)
                #print type(map_dic[type_p])
            else:
                map_dic[type_p] = [key_p]
        #print map_dic
        #print len(map_dic)
        i = 0
        for j in map_dic:
            type_dic[j] = i
            for k in map_dic[j]:
                pro_dic[k] = i
            i += 1
            #print map_dic[j]
        print type_dic
        print pro_dic


        trans = xlrd.open_workbook(tr_addr).sheet_by_index(0)
        newbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
        newsheet = newbook.add_sheet('sheet1', cell_overwrite_ok=True)
        newsheet2 = newbook.add_sheet('sheet2', cell_overwrite_ok=True)
        newsheet3 = newbook.add_sheet('sheet3', cell_overwrite_ok=True)
        newsheet4 = newbook.add_sheet('sheet4', cell_overwrite_ok=True)
        #wrow = 0
        newsheet.write(0, 0, trans.row_values(0)[0])
        newsheet.write(0, 1, trans.row_values(0)[1])
        newsheet.write(0, 2, trans.row_values(0)[2])
        for i in range(1, trans.nrows):
            newsheet.write(i, 0, trans.row_values(i)[0])
            newsheet.write(i, 1, trans.row_values(i)[1])
            j = 2
            try:
                #print trans.row_values(i)[j]
                while isinstance(trans.row_values(i)[j], float):
                    #print pro_dic[int(trans.row_values(i)[j])]
                    newsheet.write(i, j, pro_dic[int(trans.row_values(i)[j])])
                    j += 1
            except:
                pass
        '''
        map_dic = {}        #{type:[list:keys of products}
        type_dic = {}       #{type:num_order}
        pro_dic = {}        #{product:type num}
        '''
        a = 0
        for i in map_dic:
            b = 0
            newsheet2.write(a, b, i)
            b += 1
            for j in map_dic[i]:
                newsheet2.write(a, b, j)
                b += 1
            a += 1
        a = 0
        for i in type_dic:
            newsheet3.write(a, 0, i)
            newsheet3.write(a, 1, type_dic[i])
            a += 1

        a = 0
        for i in pro_dic:
            newsheet4.write(a, 0, i)
            newsheet4.write(a, 1, pro_dic[i])
            a += 1
        newbook.save("/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/processed_data/sales_m_t_m.xls")
        return 0

# PreProcess.sale_to_trans_time("/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/raw_data/sales.xlsx", "/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/processed_data/sales_mod_t.xls")
# PreProcess.sale_to_trans("/Users/hcwang/Desktop/dm/sales.xlsx", "/Users/hcwang/Desktop/dm/sales_mod.xls")
# PreProcess.mapping("/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/raw_data/product.xlsx"
#        , "/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/processed_data/sales_mod.xls")
PreProcess.mapping_t("/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/raw_data/product.xlsx"
        , "/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/processed_data/sales_mod_t.xls")


