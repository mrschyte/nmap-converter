#!env python

from libnmap.parser import NmapParser, NmapParserException
from xlsxwriter import Workbook
from datetime import datetime

import os.path

def head(xs, default=""):
    if len(xs) > 0:
        return xs[0]
    return default

def main(reports, workbook):
    summary = workbook.add_worksheet("Summary")
    results = workbook.add_worksheet("Results")
    row = 1

    for reportid, report in enumerate(reports):
        fmt_bold = workbook.add_format({"bold": True})
        fmt_conf = workbook.add_format()
        fmt_conf.set_num_format('0%')

        results.autofilter("A1:N1")
        results.freeze_panes(1, 0)

        results.data_validation("M2:M$1048576", {"validate": "list",
                                                "source": ["Y", "N", "N/A"]})

        summary_header = ["Input Name", "Command", "Version", "Scan Type", "Started", "Completed", "Hosts Total", "Hosts Up", "Hosts Down"]
        summary_body = {"Input Name": lambda report: report.basename,
                        "Command": lambda report: report.commandline,
                        "Version": lambda report: report.version,
                        "Scan Type": lambda report: report.scan_type,
                        "Started": lambda report: datetime.utcfromtimestamp(report.started).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                        "Completed": lambda report: datetime.utcfromtimestamp(report.endtime).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                        "Hosts Total": lambda report: report.hosts_total,
                        "Hosts Up": lambda report: report.hosts_up,
                        "Hosts Down": lambda report: report.hosts_down}

        results_header = ["Host", "IP", "Port", "Protocol", "Status", "Service", "Tunnel", "Method", "Confidence", "Reason", "Product", "Version", "Extra", "Flagged", "Notes"]
        results_body = {"Host": lambda host, service: head(host.hostnames),
                        "IP": lambda host, service: host.address,
                        "Port": lambda host, service: service.port,
                        "Protocol": lambda host, service: service.protocol,
                        "Status": lambda host, service: service.state,
                        "Service": lambda host, service: service.service,
                        "Tunnel": lambda host, service: service.tunnel,
                        "Method": lambda host, service: service.service_dict.get("method", ""),
                        "Confidence": lambda host, service: float(service.service_dict.get("conf", "0")) / 10,
                        "Reason": lambda host, service: service.reason,
                        "Host": lambda host, service: head(host.hostnames),
                        "Product": lambda host, service: service.service_dict.get("product", ""),
                        "Version": lambda host, service: service.service_dict.get("version", ""),
                        "Extra": lambda host, service: service.service_dict.get("extrainfo", ""),
                        "Flagged": lambda host, service: "N/A",
                        "Notes": lambda host, service: ""}

        results_format = {"Confidence": fmt_conf}

        print("[+] Processing {}".format(report.summary))
        for idx, item in enumerate(summary_header):
            summary.write(0, idx, item, fmt_bold)
            for idx, item in enumerate(summary_header):
                summary.write(1 + reportid, idx, summary_body[item](report))

        for idx, item in enumerate(results_header):
            results.write(0, idx, item, fmt_bold)

        for host in report.hosts:
            print("[+] Processing {}".format(host))
            for service in host.services:
                for idx, item in enumerate(results_header):
                    results.write(row, idx, results_body[item](host, service), results_format.get(item, None))
                row += 1

    workbook.close()
    

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("-o", "--output", metavar="XLS", help="path to xlsx output")
    parser.add_argument("reports", metavar="XML", nargs="+", help="path to nmap xml report")
    args = parser.parse_args()

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
