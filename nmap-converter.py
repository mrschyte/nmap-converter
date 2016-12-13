from libnmap.parser import NmapParser
from xlsxwriter import Workbook
from datetime import datetime

def head(xs, default=""):
    if len(xs) > 0:
        return xs[0]
    return default

def main(report, workbook):
    summary = workbook.add_worksheet("Summary")
    results = workbook.add_worksheet("Results")

    fmt_bold = workbook.add_format({"bold": True})
    fmt_conf = workbook.add_format()
    fmt_conf.set_num_format('0%')

    results.autofilter("A1:N1")
    results.freeze_panes(1, 0)

    results.data_validation("K2:K$1048576", {"validate": "list",
                                             "source": ["Y", "N", "N/A"]})

    summary_header = ["Command", "Version", "Scan Type", "Started", "Completed", "Hosts Total", "Hosts Up", "Hosts Down"]
    summary_body = {"Command": lambda report: report.commandline,
                    "Version": lambda report: report.version,
                    "Scan Type": lambda report: report.scan_type,
                    "Started": lambda report: datetime.utcfromtimestamp(report.started).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                    "Completed": lambda report: datetime.utcfromtimestamp(report.endtime).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                    "Hosts Total": lambda report: report.hosts_total,
                    "Hosts Up": lambda report: report.hosts_up,
                    "Hosts Down": lambda report: report.hosts_down}

    results_header = ["Host", "IP", "Port", "Protocol", "Status", "Service", "Method", "Confidence", "Reason", "Product", "Version", "Extra", "Flagged", "Notes"]
    results_body = {"Host": lambda host, service: head(host.hostnames),
                    "IP": lambda host, service: host.address,
                    "Port": lambda host, service: service.port,
                    "Protocol": lambda host, service: service.protocol,
                    "Status": lambda host, service: service.state,
                    "Service": lambda host, service: service.service,
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
        summary.write(idx, 0, item, fmt_bold)
        for idx, item in enumerate(summary_header):
            summary.write(idx, 1, summary_body[item](report))

    for idx, item in enumerate(results_header):
        results.write(0, idx, item, fmt_bold)

    row = 1
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
    parser.add_argument("xmlreport", help="path to nmap xml report")
    parser.add_argument("workbook", help="path to xls workbook")
    args = parser.parse_args()

    report = NmapParser.parse_fromfile(args.xmlreport)
    workbook = Workbook(args.workbook)
    main(report, workbook)
