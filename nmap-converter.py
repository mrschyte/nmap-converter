#!/usr/bin/env python

from libnmap.parser import NmapParser, NmapParserException
from xlsxwriter import Workbook
from datetime import datetime

import os.path

def generate_summary(workbook, sheet, report):
    summary_header = ["Scan", "Command", "Version", "Scan Type", "Started", "Completed", "Hosts Total", "Hosts Up", "Hosts Down"]
    summary_body = {"Scan": lambda report: report.basename,
                    "Command": lambda report: report.commandline,
                    "Version": lambda report: report.version,
                    "Scan Type": lambda report: report.scan_type,
                    "Started": lambda report: datetime.utcfromtimestamp(report.started).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                    "Completed": lambda report: datetime.utcfromtimestamp(report.endtime).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                    "Hosts Total": lambda report: report.hosts_total,
                    "Hosts Up": lambda report: report.hosts_up,
                    "Hosts Down": lambda report: report.hosts_down}

    for idx, item in enumerate(summary_header):
        sheet.write(0, idx, item, workbook.myformats["fmt_bold"])
        for idx, item in enumerate(summary_header):
            sheet.write(sheet.lastrow + 1, idx, summary_body[item](report))

    sheet.lastrow = sheet.lastrow + 1


def generate_hosts(workbook, sheet, report):
    sheet.autofilter("A1:E1")
    sheet.freeze_panes(1, 0)

    hosts_header = ["Host", "IP", "Status", "Services", "OS"]
    hosts_body = {"Host": lambda host: next(iter(host.hostnames), ""),
                  "IP": lambda host: host.address,
                  "Status": lambda host: host.status,
                  "Services": lambda host: len(host.services),
                  "OS": lambda host: os_class_string(host.os_class_probabilities())}

    for idx, item in enumerate(hosts_header):
        sheet.write(0, idx, item, workbook.myformats["fmt_bold"])

    row = sheet.lastrow
    for host in report.hosts:
        for idx, item in enumerate(hosts_header):
            sheet.write(row + 1, idx, hosts_body[item](host))
        row += 1

    sheet.lastrow = row

def generate_results(workbook, sheet, report):
    sheet.autofilter("A1:N1")
    sheet.freeze_panes(1, 0)

    sheet.data_validation("N2:N$1048576", {"validate": "list",
                                           "source": ["Y", "N", "N/A"]})

    results_header = ["Host", "IP", "Port", "Protocol", "Status", "Service", "Tunnel", "Method", "Confidence", "Reason", "Product", "Version", "Extra", "Flagged", "Notes"]
    results_body = {"Host": lambda host, service: next(iter(host.hostnames), ""),
                    "IP": lambda host, service: host.address,
                    "Port": lambda host, service: service.port,
                    "Protocol": lambda host, service: service.protocol,
                    "Status": lambda host, service: service.state,
                    "Service": lambda host, service: service.service,
                    "Tunnel": lambda host, service: service.tunnel,
                    "Method": lambda host, service: service.service_dict.get("method", ""),
                    "Confidence": lambda host, service: float(service.service_dict.get("conf", "0")) / 10,
                    "Reason": lambda host, service: service.reason,
                    "Product": lambda host, service: service.service_dict.get("product", ""),
                    "Version": lambda host, service: service.service_dict.get("version", ""),
                    "Extra": lambda host, service: service.service_dict.get("extrainfo", ""),
                    "Flagged": lambda host, service: "N/A",
                    "Notes": lambda host, service: ""}

    results_format = {"Confidence": workbook.myformats["fmt_conf"]}

    print("[+] Processing {}".format(report.summary))
    for idx, item in enumerate(results_header):
        sheet.write(0, idx, item, workbook.myformats["fmt_bold"])

    row = sheet.lastrow
    for host in report.hosts:
        print("[+] Processing {}".format(host))
        for service in host.services:
            for idx, item in enumerate(results_header):
                sheet.write(row + 1, idx, results_body[item](host, service), results_format.get(item, None))
            row += 1

    sheet.lastrow = row

def setup_workbook_formats(workbook):
    formats = {"fmt_bold": workbook.add_format({"bold": True}),
               "fmt_conf": workbook.add_format()}

    formats["fmt_conf"].set_num_format("0%")
    return formats

def os_class_string(os_class_array):
    return " | ".join(["{0} ({1}%)".format(os_string(osc), osc.accuracy) for osc in os_class_array])

def os_string(os_class):
    rval = "{0}, {1}".format(os_class.vendor, os_class.osfamily)
    if len(os_class.osgen):
        rval += "({0})".format(os_class.osgen)
    return rval

def main(reports, workbook):
    sheets = {"Summary": generate_summary,
              "Hosts": generate_hosts,
              "Results": generate_results}

    workbook.myformats = setup_workbook_formats(workbook)

    for sheet_name, sheet_func in sheets.items():
        sheet = workbook.add_worksheet(sheet_name)
        sheet.lastrow = 0
        for report in reports:
            sheet_func(workbook, sheet, report)
    workbook.close()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("-o", "--output", metavar="XLS", help="path to xlsx output")
    parser.add_argument("reports", metavar="XML", nargs="+", help="path to nmap xml report")
    args = parser.parse_args()

    if args.output == None:
        parser.error("Output must be specified")

    reports = []
    for report in args.reports:
        try:
            parsed = NmapParser.parse_fromfile(report)
        except NmapParserException as ex:
            parsed = NmapParser.parse_fromfile(report, incomplete=True)
        
        parsed.basename = os.path.basename(report)
        reports.append(parsed)

    workbook = Workbook(args.output)
    main(reports, workbook)
