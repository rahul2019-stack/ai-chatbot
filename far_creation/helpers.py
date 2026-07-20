import csv
import pandas as pd
from config import config_dict

def create_excel_sheet(env, app, department, bastion_node_ips, bootstrap_node_ips, master_node_ips, worker_node_ips, infra_node_ips, vip1, vip2):
    description_df = _form_description_data_df(env, app, department, bastion_node_ips, bootstrap_node_ips, master_node_ips, worker_node_ips, infra_node_ips, vip1, vip2)
    lb_df = _form_lb_data_df(env, app, department)
    with pd.ExcelWriter("far.xlsx") as writer:
        description_df.to_excel(writer, sheet_name="Description", index=False)

def _form_lb_data_df(env, app, department):
    data = {
        "FrontEnd": [f"api.{app}{env}.bank.sbi", f"api-int.{app}{env}.bank.sbi", f"*.apps.{app}{env}.bank.sbi"],
        "FrontEndPort": [6443, ],
        "Backend": ["Load Balancer 1", "Load Balancer 2"],
        "BackendPort": []
    }
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