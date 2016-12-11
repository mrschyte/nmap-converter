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

    bold = workbook.add_format({"bold": True})
    flag = workbook.add_format({"bg_color":   "#FFC7CE",
                                "font_color": "#9C0006"})
    results.autofilter("A1:L1")
    results.freeze_panes(1, 0)

    results.data_validation("K2:K$1048576", {"validate": "list",
                                             "source": ["Y", "N"]})

    summary_header = ["Command", "Version", "Scan Type", "Started", "Completed", "Hosts Total", "Hosts Up", "Hosts Down"]
    summary_body = {"Command": lambda report: report.commandline,
                    "Version": lambda report: report.version,
                    "Scan Type": lambda report: report.scan_type,
                    "Started": lambda report: datetime.utcfromtimestamp(report.started).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                    "Completed": lambda report: datetime.utcfromtimestamp(report.endtime).strftime("%Y-%m-%d %H:%M:%S (UTC)"),
                    "Hosts Total": lambda report: report.hosts_total,
                    "Hosts Up": lambda report: report.hosts_up,
                    "Hosts Down": lambda report: report.hosts_down}

    results_header = ["Host", "IP", "Port", "Protocol", "Status", "Service", "Reason", "Product", "Version", "Extra", "Flagged", "Notes"]
    results_body = {"Host": lambda host, port, service: head(host.hostnames),
                    "IP": lambda host, port, service: host.address,
                    "Port": lambda host, port, service: port[0],
                    "Protocol": lambda host, port, service: port[1],
                    "Status": lambda host, port, service: service.state,
                    "Service": lambda host, port, service: service.service,
                    "Reason": lambda host, port, service: service.reason,
                    "Host": lambda host, port, service: head(host.hostnames),
                    "Product": lambda host, port, service: service.service_dict.get("product", ""),
                    "Version": lambda host, port, service: service.service_dict.get("version", ""),
                    "Extra": lambda host, port, service: service.service_dict.get("extrainfo", ""),
                    "Flagged": lambda host, port, service: "N",
                    "Notes": lambda host, port, service: ""}

    print("[+] Processing {}".format(report.summary))
    for idx, item in enumerate(summary_header):
        summary.write(idx, 0, item, bold)
        for idx, item in enumerate(summary_header):
            summary.write(idx, 1, summary_body[item](report))

    for idx, item in enumerate(results_header):
        results.write(0, idx, item, bold)

    row = 1
    for host in report.hosts:
        print("[+] Processing {}".format(host))
        for port in host.get_ports():
            for idx, item in enumerate(results_header):
                service = host.get_service(port[0], port[1])
                results.write(row, idx, results_body[item](host, port, service))
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
