#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import re
from pyzabbix import ZabbixAPI

if __name__ == "__main__":
    url = 'http://192.168.56.8/zabbix'
    login_id = 'admin'
    password = 'zabbix'

    api = ZabbixAPI(url)
    api.login(login_id, password)

    book = xlrd.open_workbook('zabbix_hearing_sheet.xlsx')
    sheet_1 = book.sheet_by_index(0)

    for row in range(sheet_1.nrows):
        resource = {}
        for col in range(sheet_1.ncols):
            cell_val = sheet_1.cell(row, col).value
            if col == 1:
                regexp = re.compile(r'^[0-9A-Za-z]+$')
                if regexp.search(cell_val) is None:
                    break
                resource['hostname'] = cell_val
            elif col == 2:
                resource['ip'] = cell_val
            elif col == 3:
                resource['os'] = cell_val
            elif col == 4:
                resource['host_group'] = cell_val
            elif col == 5:
                resource['ping_interval'] = cell_val
            elif col == 6:
                resource['detection_count'] = cell_val

        if not len(resource):
            continue

        hg_filter = {'name': resource['host_group']}
        host_group = api.hostgroup.get(output='extend', filter=hg_filter)[0]
        if not host_group:
            api.hostgroup.create({'name': resource['host_group']})

        h_filter = {'host': resource['hostname']}
        host = api.host.get(output='extend', filter=h_filter)

        if not host:
            param = {}
            group_list = []
            groups = {}
            groups['groupid'] = host_group['groupid']
            group_list.append(groups)
            param['host'] = resource['hostname']
            param['ip'] = resource['ip']
            param['port'] = 10050
            param['useip'] = 1
            param['groups'] = group_list
            result = api.host.create(param)

            param = {}
            param['description'] = 'ICMP ping'
            param['key_'] = 'icmpping'
            param['hostid'] = result['hostids'][0]
            param['delay'] = resource['ping_interval']
            api.item.create(param)
