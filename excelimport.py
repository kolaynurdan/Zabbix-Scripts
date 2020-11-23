# -*- coding: utf-8 -*- 
import xlrd
from pyzabbix import ZabbixAPI
import pprint

import sys
reload(sys)
sys.setdefaultencoding( "utf-8" )

#
file_name=raw_input('path:')

#
try:
    zapi = ZabbixAPI("https://Zabbix_Url")
    zapi.login("Zabbix_user", "Zabbix_password")
except:
    print '\n hata olustu'
    sys.exit()

#
def get_templateid(template_name):
    template_data = {
        "host": [template_name]
    }
    result = zapi.template.get(filter=template_data)
    if result:
        return result[0]['templateid']
    else:
        return result

#
def check_group(group_name):
    return zapi.hostgroup.get(output= "groupid",filter= {"name": group_name})

#
def create_group(group_name):
    groupid=zapi.hostgroup.create({"name": group_name})

#
def get_groupid(group_name):
    group_id=check_group(group_name)[0]["groupid"]
    return group_id

#
def create_host(host_data):
    host=zapi.host.get(output= ["host"],filter= {"host": host_data["host"]})
    if len(host) > 0 and host[0]["host"] == host_data["host"]:
      #print "host %s exists" % host_data["name"]
      print "host exists: %s ,group: %s ,templateid: %s" % (host_data["name"],host_data["groups"][0]["groupid"],host_data["templates"][0]["templateid"])
    else:
      zapi.host.create(host_data)
      print "host created: %s ,group: %s ,templateid: %s" % (host_data["name"],host_data["groups"][0]["groupid"],host_data["templates"][0]["templateid"])

#
def open_excel(file= file_name):
     try:
         data = xlrd.open_workbook(file)
         return data
     except Exception,e:
         print str(e)

#
def get_hosts(file):
    data = open_excel(file)
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    list = []
    for rownum in range(1,nrows):
      #print table.row_values(rownum)[0]
      list.append(table.row_values(rownum)) 
    return list

def main():
  hosts=get_hosts(file_name)
  for host in hosts:
    host_name=host[6]
    visible_name="Dogus "+host[1]+" "+host[5]+" | "+host[6]
    host_ip=host[6]
    group=host[0]+"/"+host[2]+"/"+host[3]
    polling_method =host[7]
    if polling_method == "SNMP":
        template = "Template Net Network Generic Device SNMPv2"
    else:
        template = "Template Module ICMP Ping"
    templateid=get_templateid(template)
    #inventory_location=host[5]
    #print templateid
    if not check_group(group):
        print u'Added host grup: %s' % group
        groupid=create_group(group)
        #print groupid
    groupid=get_groupid(group)
    host_data = {
        "host": host_name,
        "name": visible_name,
        "interfaces": [
            {
                "type": 2,
                "main": 1,
                "useip": 1,
                "ip": host_ip.strip(),
                "dns": "",
                "port": "161",
                "details": {
                    "version": 2,
                    "community": "public"
                }
            }
        ],
        "groups": [
            {
                "groupid": groupid
            }
        ],
        "templates": [
            {
                "templateid": templateid
            }
        ]
    }
    #print "visiblename: %s ,group: %s ,templateid: %s" % (visible_name,group,templateid)
    create_host(host_data)

if __name__=="__main__":
    main()