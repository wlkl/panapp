#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
__author__ = 'wlkl'

import sys
import os
import telnetlib
import argparse
import re
import xlrd
from socket import *
import ipaddress
import time

if os.geteuid() != 0:
    print("\033[1;31mДля запуска этого скрипта нужно быть Рутом!\033[1;m")
    pass
    #sys.exit()

def xlsfile(w_file):
    global wb
    try:
        wb = xlrd.open_workbook(w_file)
    except (xlrd.mmap.error, xlrd.biffh.XLRDError):
        print(w_file, ' не является файлом MS Excel!')
        exit(1)
    except IOError:
        print(w_file, ' не существует, или ошибка в имяни')
        exit(1)
    else:
        return wb

parser = argparse.ArgumentParser(description='Mass PM processing')
parser.add_argument('-n', '--net', dest='net', metavar='Net', help='Net in X.X.X.X/X form', required=True)
parser.add_argument('-f', '--file', dest='work_file', metavar='File', type=xlsfile, help='.xls file with data', required=True)
parser.add_argument('-s', '--sheet', dest='sheet_index', metavar='Sheet', type=int, help='Type here sheet number that you want to use, by default it will be 1', default=1, required=False)
args = parser.parse_args()

work_net = ipaddress.ip_network(args.net)
sheet = wb.sheet_by_index(args.sheet_index - 1)
sheet_name = str(sheet.name)
port_for_scan = 23
#f = open("pan_log.txt", "w")

def telnet_hosts(net):
    pim_ip = []
    for host in net.hosts():
        s = socket(AF_INET, SOCK_STREAM)
        s.settimeout(0.1)
        res = s.connect_ex((str(host), 23))
        if res == 0:
            pim_ip.append(str(host))
    print("PM's IP: %s \n" % pim_ip)
    return pim_ip

def ret_mac_xls(mac_addr):
    ret = {}
    for row in range(sheet.nrows):
        ro = sheet.cell(row, 13).value
        if re.search(mac_addr, ro):
            ret.update({'rackposition':int(sheet.cell(row, 10).value)})
            ret.update({'rackname':str(sheet.cell(row, 7).value)})
            ret.update({'dev_name':str(sheet.cell(row, 9).value)})
            ret.update({'room':int(sheet.cell(row, 6).value)})
            ret.update({'ip':str(sheet.cell(row, 14).value)})
            ret.update({'mask':str(sheet.cell(row, 18).value)})
            ret.update({'gw':str(sheet.cell(row, 19).value)})
            ret.update({'trapip':str(sheet.cell(row, 16).value)})
            row += 1
            offset = 2
            em = 0
            c = True
            while c:
                if re.search('-EM$', sheet.cell(row, 12).value):
                    ret.update({"offsetname_%s" % offset:str(sheet.cell(row, 9).value)})
                    ret.update({"offset_rackposition_%s" % offset:int(sheet.cell(row, 10).value)})
                    ret.update({"offset_rackname_%s" % offset:str(sheet.cell(row, 7).value)})
                    row += 1
                    em += 1
                    offset += 1
                else:
                    c = False
            ret.update({'em': em})
            return ret
    if not ret:
        return None

for ho in telnet_hosts(work_net):
    tn = telnetlib.Telnet(ho)
    print(tn.read_some())
    if re.search("Welcome to Panduit Shell", str(tn.read_some())):
        pass
    else:
        print("%s не PM! пропускаем..." % ho)
        #f.write("%s не PM! пропускаем...\n" % ho)
        tn.close()
        continue
    tn.write("admin\n")
    tn.write("panduit\n")
    tn.read_until("PSH >")
    tn.write("show mac\n")
    mac = re.search(r'..:..:..:..:..:..', tn.read_until("PSH >")).group()
    val = ret_mac_xls(mac)
    if val == None:
        tn.close()
        print("MAC %s is not found!" % mac)
        #f.write("MAC %s is not found!\n" % mac)
        continue
    file_log = open(val.get("dev_name")+".txt", "w")
    tn.write("config devicename %s\n" % val.get("dev_name").encode('ascii'))
    file_log.write(tn.read_until("PSH >"))
    tn.write("config physloc %s/%s\n" % (sheet_name.encode('ascii'), val.get("room")))
    file_log.write(tn.read_until("PSH >"))
    tn.write("config rackposition 1 %s\n" % val.get("rackposition"))
    file_log.write(tn.read_until("PSH >"))
    tn.write("config rackname 1 %s\n" % val.get("rackname").encode('ascii'))
    file_log.write(tn.read_until("PSH >"))
    if val.get("em"):
        off = 2
        for e in range(val.get("em")):
            tn.write("config offsetname %s %s\n" % (off, val.get("offsetname_%s" % off).encode('ascii')))
            file_log.write(tn.read_until("PSH >"))
            tn.write("config rackposition %s %s\n" % (off, val.get("offset_rackposition_%s" % off)))
            file_log.write(tn.read_until("PSH >"))
            tn.write("config rackname %s %s\n" % (off, val.get("offset_rackname_%s" % off).encode('ascii')))
            file_log.write(tn.read_until("PSH >"))
            off += 1
    tn.write("config snmp -trapip 1 %s\n" % val.get("trapip").encode('ascii'))
    file_log.write(tn.read_until("PSH >"))
    tn.write("config snmp -trapon 1 all\n")
    file_log.write(tn.read_until("PSH >"))
    tn.write("config ip -type static -addr %s -mask %s -gateway %s\n" % (val.get("ip").encode('ascii'), val.get("mask").encode('ascii'), val.get("gw").encode('ascii')))
    file_log.write(tn.read_until("PSH >"))
    tn.write("config ip commit\n")
    file_log.write(tn.read_until("PSH >"))
    tn.write("config ip config\n")
    file_log.write(tn.read_until("PAN SHELL is closed by your end, bye ..") + "\n")
    tn.close()
    print("PM с MAC адресом %s обработан\n" % mac)
    file_log.close()
#f.close()
print("ВСЁ!")
exit()