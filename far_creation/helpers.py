import csv
import pandas as pd
from config import config_dict
from openpyxl import load_workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Side,
    Font
)
from openpyxl.utils import get_column_letter

def create_excel_sheet(env, app, department, bastion_node_ips, bootstrap_node_ips, master_node_ips, worker_node_ips, infra_node_ips, vip1, vip2):
    description_df = _form_description_data_df(env, app, department, bastion_node_ips, bootstrap_node_ips, master_node_ips, worker_node_ips, infra_node_ips, vip1, vip2)
    lb_df = _form_lb_data_df(env, app, department, bootstrap_node_ips, master_node_ips, infra_node_ips)
    with pd.ExcelWriter("far.xlsx") as writer:
        description_df.to_excel(writer, sheet_name="Description", index=False)
        lb_df.to_excel(writer, sheet_name="LB_Details", index=False)
    _format_sheet(
        excel_file="far.xlsx",
        sheet_name="LB_Details",
        merge_columns=["A"]
    )
    _format_sheet(
        excel_file="far.xlsx",
        sheet_name="Description"
    )

def _format_sheet(
    excel_file,
    sheet_name,
    merge_columns=None
):
    """
    Formats an Excel worksheet.

    Args:
        excel_file (str): Path to Excel file.
        sheet_name (str): Name of worksheet.
        merge_columns (list): List of Excel column letters to merge.
                              Example: ["A", "B"]
    """

    wb = load_workbook(excel_file)
    ws = wb[sheet_name]

    # -------------------------
    # Header Formatting
    # -------------------------

    header_font = Font(bold=True)

    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    center_alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True    
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center_alignment

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border

    # -------------------------
    # Merge Duplicate Cells
    # -------------------------

    if merge_columns:

        for column in merge_columns:

            start_row = 2

            while start_row <= ws.max_row:

                value = ws[f"{column}{start_row}"].value
                end_row = start_row

                while (
                    end_row < ws.max_row
                    and ws[f"{column}{end_row+1}"].value == value
                ):
                    end_row += 1

                if end_row > start_row:
                    ws.merge_cells(
                        f"{column}{start_row}:{column}{end_row}"
                    )

                    ws[f"{column}{start_row}"].alignment = center_alignment

                start_row = end_row + 1

    # -------------------------
    # Auto Column Width
    # -------------------------

    for column in ws.columns:

        max_length = 0
        column_letter = get_column_letter(column[0].column)

        for cell in column:
            try:
                if cell.value:
                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )
            except:
                pass

        ws.column_dimensions[column_letter].width = max_length + 3

    wb.save(excel_file)

def _form_lb_data_df(env, app, department, bootstrap_node_ips, master_node_ips, infra_node_ips):
    total_bootstrap_and_master_nodes = len(bootstrap_node_ips) + len(master_node_ips)
    total_infra_nodes = len(infra_node_ips)
    total_master_nodes = len(master_node_ips)
    api_server = f"api.{app}{env}.bank.sbi"
    api_int_server = f"api-int.{app}{env}.bank.sbi"
    apps_server = f"*.apps.{app}{env}.bank.sbi"

    bootstrap_dns = f"bootstrap.{app}{env}.bank.sbi"
    master_dns = f"master.{app}{env}.bank.sbi"
    infra_dns = f"infra.{app}{env}.bank.sbi"

    front_back_end_port = []
    front_back_end_port.extend(['6443'] * total_bootstrap_and_master_nodes)
    front_back_end_port.extend(['22623'] * total_bootstrap_and_master_nodes)
    front_back_end_port.extend(['443'] * total_infra_nodes)

    front_end_list = []
    front_end_list.extend([api_server] * total_bootstrap_and_master_nodes)
    front_end_list.extend([api_int_server] * total_bootstrap_and_master_nodes)
    front_end_list.extend([apps_server] * total_infra_nodes)

    # backend_dns = [bootstrap_dns, master_dns * total_master_nodes, bootstrap_dns, master_dns * total_master_nodes, infra_dns * total_infra_nodes]
    backend_dns = []
    backend_dns.extend([bootstrap_dns])
    backend_dns.extend([master_dns] * total_master_nodes)
    backend_dns.extend([bootstrap_dns])
    backend_dns.extend([master_dns] * total_master_nodes)
    backend_dns.extend([infra_dns] * total_infra_nodes)

    data = {
        "FrontEnd": front_end_list,
        "FrontEndPort": front_back_end_port,
        "Backend": backend_dns,
        "BackendPort": front_back_end_port,
    }

    print(f"Data formed for LB details is {data}")
    df = pd.DataFrame(data)
    return df

def _form_description_data_df(env, app, department, bastion_node_ips, bootstrap_node_ips, master_node_ips, worker_node_ips, infra_node_ips, vip1, vip2):
    node_ips = ""
    for i, node_ip in enumerate(bastion_node_ips, start=1):
        if len(bastion_node_ips) == 1:
            node_ips += f" Bastion/Registry {node_ip}\n"
        else:
            node_ips = f" Registry {i} {node_ip}\n"
    for i, node_ip in enumerate(bootstrap_node_ips, start=1):
        if len(bootstrap_node_ips) == 1:
            node_ips += f" Bootstrap {node_ip}\n"
        else:
            node_ips += f" Bootstrap {i} {node_ip}\n"
    for i, node_ip in enumerate(master_node_ips, start=1):
        if len(master_node_ips) == 1:
            node_ips += f" Master {node_ip}\n"
        else:
            node_ips += f" Master {i} {node_ip}\n"
    for i, node_ip in enumerate(worker_node_ips, start=1):
        if len(worker_node_ips) == 1:
            node_ips += f" Worker {node_ip}\n"
        else:
            node_ips += f" Worker {i} {node_ip}\n"
    if len(infra_node_ips) > 0:
        for i, node_ip in enumerate(infra_node_ips, start=1):
            if len(infra_node_ips) == 1:
                node_ips += f" Infra {node_ip}\n"
            else:
                node_ips += f" Infra {i} {node_ip}\n"

    if env.lower() == "uat" or env.lower() == "pre-prod" or env.lower() == "dev":
        node_ips += f" Central Bastion {config_dict['central_meghdoot_bastion_node_ip_uat']}\n"
        node_ips += f" Central Mirror {config_dict['central_meghdoot_mirror_node_ip_uat']}\n"
    else:
        node_ips += f" Central Bastion {config_dict['central_meghdoot_bastion_node_ip_prod']}\n"
        node_ips += f" Central Mirror {config_dict['central_meghdoot_mirror_node_ip_prod']}\n"
    print(f"Node  string formed is {node_ips}")
    vip_str = f"VIP1: {vip1}\nVIP2: {vip2}"
    ldap_ip_str = "\n".join(config_dict["LDAP_ip"])
    ntp_ip_str = "\n".join(config_dict["NTP_ip"])
    ocp_team_desktop_ip_str = "\n".join(config_dict["ocp_team_desktop_ip"])

    data = {
        "Host-Description": ["Node Info", "LB IP", "LDAP url", "LDAP IP", "NTP IP", "Windows IP"],
        "Host-Ips": [node_ips, vip_str, config_dict["LDAP_url"], ldap_ip_str, ntp_ip_str, ocp_team_desktop_ip_str],
        "Details": ["Node Ip info", "VIP info", "This url is used to establish connection between Ocp cluster and vcenter platform as well as authentication",
                    "Used for authentication", "Used for time sync", "Meghdoot OCP Team's desktop IP"]
    }
    df = pd.DataFrame(data)
    return df