#!/usr/bin/env python

from libnmap.parser import NmapParser, NmapParserException
from xlsxwriter import Workbook
from datetime import datetime

import os.path


class HostModule():
    def __init__(self, host):
        self.host = next(iter(host.hostnames), "")
        self.ip = host.address
        self.port = ""
        self.protocol = ""
        self.status = ""
        self.service = ""
        self.tunnel = ""
        self.method = ""
        self.source = ""
        self.confidence = ""
        self.reason = ""
        self.reason = ""
        self.product = ""
        self.version = ""
        self.extra = ""
        self.flagged = "N/A"
        self.notes = ""


class ServiceModule(HostModule):
    def __init__(self, host, service):
        super(ServiceModule, self).__init__(host)
        self.host = next(iter(host.hostnames), "")
        self.ip = host.address
        self.port = service.port
        self.protocol = service.protocol
        self.status = service.state
        self.service = service.service
        self.tunnel = service.tunnel
        self.method = service.service_dict.get("method", "")
        self.source = "scanner"
        self.confidence = float(service.service_dict.get("conf", "0")) / 10
        self.reason = service.reason
        self.product = service.service_dict.get("product", "")
        self.version = service.service_dict.get("version", "")
        self.extra = service.service_dict.get("extrainfo", "")
        self.flagged = "N/A"
        self.notes = ""


class HostScriptModule(HostModule):
    def __init__(self, host, script):
        super(HostScriptModule, self).__init__(host)
        self.method = script["id"]
        self.source = "script"
        self.extra = script["output"].strip()


class ServiceScriptModule(ServiceModule):
    def __init__(self, host, service, script):
        super(ServiceScriptModule, self).__init__(host, service)
        self.source = "script"
        self.method = script["id"]
        self.extra = script["output"].strip()


def _tgetattr(object, name, default=None):
    try:
        return getattr(object, name, default)
    except Exception:
        return default


def generate_summary(workbook, sheet, report):
    summary_header = ["Scan", "Command", "Version", "Scan Type", "Started", "Completed", "Hosts Total", "Hosts Up", "Hosts Down"]
    summary_body = {"Scan": lambda report: _tgetattr(report, 'basename', 'N/A'),
                    "Command": lambda report: _tgetattr(report, 'commandline', 'N/A'),
                    "Version": lambda report: _tgetattr(report, 'version', 'N/A'),
                    "Scan Type": lambda report: _tgetattr(report, 'scan_type', 'N/A'),
                    "Started": lambda report: datetime.utcfromtimestamp(_tgetattr(report, 'started', 0)).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                    "Completed": lambda report: datetime.utcfromtimestamp(_tgetattr(report, 'endtime', 0)).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                    "Hosts Total": lambda report: _tgetattr(report, 'hosts_total', 'N/A'),
                    "Hosts Up": lambda report: _tgetattr(report, 'hosts_up', 'N/A'),
                    "Hosts Down": lambda report: _tgetattr(report, 'hosts_down', 'N/A')}

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

    results_header = ["Host", "IP", "Port", "Protocol", "Status", "Service", "Tunnel", "Source", "Method", "Confidence", "Reason", "Product", "Version", "Extra", "Flagged", "Notes"]
    results_body = {"Host": lambda module: module.host,
                    "IP": lambda module: module.ip,
                    "Port": lambda module: module.port,
                    "Protocol": lambda module: module.protocol,
                    "Status": lambda module: module.status,
                    "Service": lambda module: module.service,
                    "Tunnel": lambda module: module.tunnel,
                    "Source": lambda module: module.source,
                    "Method": lambda module: module.method,
                    "Confidence": lambda module: module.confidence,
                    "Reason": lambda module: module.reason,
                    "Product": lambda module: module.product,
                    "Version": lambda module: module.version,
                    "Extra": lambda module: module.extra,
                    "Flagged": lambda module: module.flagged,
                    "Notes": lambda module: module.notes}

    results_format = {"Confidence": workbook.myformats["fmt_conf"]}

    print("[+] Processing {}".format(report.summary))
    for idx, item in enumerate(results_header):
        sheet.write(0, idx, item, workbook.myformats["fmt_bold"])

    row = sheet.lastrow
    for host in report.hosts:
        print("[+] Processing {}".format(host))

        for script in host.scripts_results:
            module = HostScriptModule(host, script)
            for idx, item in enumerate(results_header):
                sheet.write(row + 1, idx, results_body[item](module), results_format.get(item, None))
            row += 1

        for service in host.services:
            module = ServiceModule(host, service)
            for idx, item in enumerate(results_header):
                sheet.write(row + 1, idx, results_body[item](module), results_format.get(item, None))
            row += 1

            for script in service.scripts_results:
                module = ServiceScriptModule(host, service, script)
                for idx, item in enumerate(results_header):
                    sheet.write(row + 1, idx, results_body[item](module), results_format.get(item, None))
                row += 1

    sheet.data_validation("O2:O${}".format(row + 1), {"validate": "list",
                                           "source": ["Y", "N", "N/A"]})
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
