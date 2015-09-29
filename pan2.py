#!/usr/bin/env python3

import xlrd
import argparse
import re
import subprocess as sp
import telnetlib
import string
import configparser
from socket import *
import ipaddress

def xlsfile(w_file):
    global wb
    try:
        wb = xlrd.open_workbook(w_file)
    except xlrd.mmap.error:
        print("%s не является файлом в формате Excel" % w_file, end="\n")
        exit(0)
    except IOError:
        print("%s не существует, или ошибка в пути к файлу" % w_file, end="\n")
        exit(0)
    else:
        return wb

parser = argparse.ArgumentParser(description="Утилита для программирования PM")
parser.add_argument("-f", "--file", dest="work_file", metavar="File", type=xlsfile, help="Файл с данными в формате Excel", required=True)
parser.add_argument("-s", "--sheet", dest="sheet_index", metavar="Sheet", type=int, help="Номер вкладки из которой брать информацию, по умолчанию 1", default=1, required=False)
args = parser.parse_args()
config = configparser.ConfigParser()
sheet = wb.sheet_by_index(args.sheet_index - 1)
#sheet_name = str(sheet.name)

config.read(r'settings.conf')

def get_mac():
    """try:
        stdout,stderr = sp.Popen(["pviqutil.exe"], stdout=sp.PIPE, stderr=sp.PIPE, stdin=sp.PIPE).communicate(b'\n')
    except WindowsError:
        print("Не могу найти файл pviqutil.exe,\n убедитесь, что он находится в каталоге указанном в переменной окружения PATH.")
        exit(1)"""
    ipnmac = {}
    """for line in stdout.decode().split('\n'):
        search_line_mac = re.search(r"([0-9A-F]{2}[:-]){5}([0-9A-F]{2})", line)
        search_line_ip = re.search(r"((2[0-5]|1[0-9]|[0-9])?[0-9]\.){3}((2[0-5]|1[0-9]|[0-9])?[0-9])", line)
        if search_line_mac:
            ipnmac.update({search_line_mac.group():search_line_ip.group()})"""

    return ipnmac

def get_data_xls(mac_addr):
    ret = {}
    coll_index = [u for u in string.ascii_lowercase]
    coll_index2 = [leg + beg for leg in string.ascii_lowercase for beg in string.ascii_lowercase]
    coll_index += coll_index2
    for row in range(sheet.nrows):
        ro = sheet.cell(row, coll_index.index(config['pim']['macaddress'])).value
        if re.search(mac_addr, ro):
            ret.update({'rackposition': int(sheet.cell(row, coll_index.index(config['pim']['rackposition'])).value)})
            ret.update({'rackname': str(sheet.cell(row, coll_index.index(config['pim']['rackname'])).value)})
            ret.update({'devicename': str(sheet.cell(row, coll_index.index(config['pim']['devicename'])).value)})
            if type(sheet.cell(row, coll_index.index(config['pim']['physloc'])).value) == float:
                ret.update({'physloc': int(sheet.cell(row, coll_index.index(config['pim']['physloc'])).value)})
            else:
                ret.update({'physloc': str(sheet.cell(row, coll_index.index(config['pim']['physloc'])).value)})
            ret.update({'ipaddress': str(sheet.cell(row, coll_index.index(config['pim']['ipaddress'])).value)})
            ret.update({'mask': str(sheet.cell(row, coll_index.index(config['pim']['mask'])).value)})
            ret.update({'gateway': str(sheet.cell(row, coll_index.index(config['pim']['gateway'])).value)})
            ret.update({'snmp_ptrapip': str(sheet.cell(row, coll_index.index(config['pim']['snmp_ptrapip'])).value)})
            row += 1
            offset = 2
            em = 0
            c = True
            while c:
                if re.search(r'^$', sheet.cell(row, coll_index.index(config['pim']['macaddress'])).value) and sheet.cell(row, coll_index.index(config['pim']['devicename'])).value:
                    ret.update({"offsetname_{}".format(offset): str(sheet.cell(row, coll_index.index(config['pim']['devicename'])).value)})
                    ret.update({"offset_rackposition_{}".format(offset): int(sheet.cell(row, coll_index.index(config['pim']['rackposition'])).value)})
                    ret.update({"offset_rackname_{}".format(offset): str(sheet.cell(row, coll_index.index(config['pim']['rackname'])).value)})
                    row += 1
                    em += 1
                    offset += 1
                else:
                    c = False
            ret.update({"em": em})
            return ret
    if not ret:
        return None

def write(by_str, file):
    file.write(by_str.decode("utf-8"))

def conf_pm(ip, data):
    log = open(data.get("devicename")+".txt", "a")
    try:
        telnet = telnetlib.Telnet(ip)
    except socket.error as err:
        log.write(err)
        log.close()
        return err
    telnet.write(b"admin\n")
    telnet.write(b"panduit\n")
    write(telnet.read_until(b"PSH >"), log)
    telnet.write("config devicename {!s}\n".format(data.get("devicename")).encode("ascii"))
    write(telnet.read_until(b"PSH >"), log)
    telnet.write("config physloc {!s}\n".format(data.get("physloc")).encode("ascii"))
    write(telnet.read_until(b"PSH >"), log)
    telnet.write("config rackposition 1 {}\n".format(data.get("rackposition")).encode("ascii"))
    write(telnet.read_until(b"PSH >"), log)
    telnet.write("config rackname 1 {!s}\n".format(data.get("rackname")).encode("ascii"))
    write(telnet.read_until(b"PSH >"), log)
    if data.get("em"):
        off = 2
        for e in range(data.get("em")):
            telnet.write("config offsetname {} {!s}\n".format(off, data.get("offsetname_{}".format(off))).encode("ascii"))
            write(telnet.read_until(b"PSH >"), log)
            telnet.write("config rackposition {} {!s}\n".format(off, data.get("offset_rackposition_{}".format(off))).encode("ascii"))
            write(telnet.read_until(b"PSH >"), log)
            telnet.write("config rackname {} {!s}\n".format(off, data.get("offset_rackname_{}".format(off))).encode("ascii"))
            write(telnet.read_until(b"PSH >"), log)
            off += 1
    telnet.write("config snmp -trapip 1 {!s}\n".format(data.get("snmp_ptrapip")).encode("ascii"))
    write(telnet.read_until(b"PSH >"), log)
    telnet.write("config snmp -trapon 1 all\n".encode("ascii"))
    write(telnet.read_until(b"PSH >"), log)
    telnet.write("config ip -type static -addr {!s} -mask {!s} -gateway {!s}\n".format(data.get("ipaddress"), data.get("mask"), data.get("gateway")).encode("ascii"))
    write(telnet.read_until(b"PSH >"), log)
    telnet.write("config ip commit\n".encode("ascii"))
    write(telnet.read_until(b"PSH >"), log)
    telnet.write("config ip config\n".encode("ascii"))
    write(telnet.read_until(b"PAN SHELL is closed by your end, bye .."), log)
    telnet.close()
    log.close()
    return 0

def main():
    mac_ip = get_mac()
    if mac_ip:
        print("Найдено {} устройств.".format(len(mac_ip)))
        for pm in mac_ip:
            print("Обрабатываю ", pm, " ... ", end="")
            xls_data = get_data_xls(pm)
            if xls_data:
                if conf_pm(mac_ip[pm], get_data_xls(pm)) == 0:
                    print("Готово!")
                else:
                    print("Устройство не отвечает, пропускаю...")
            else:
                print("В файле Excel, закладка {!s} отсутствует запись об устройстве с таким МАК-адресом.".format(sheet.name))
                continue
    else:
        print("В сети не найдено ни одного PM!")
        exit(0)

if __name__ == '__main__':
    main()

