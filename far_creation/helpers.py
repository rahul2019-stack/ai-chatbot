import csv
import pandas as pd

def create_excel_sheet(env, app, department, bastion_node_ips, bootstrap_node_ips, master_node_ips, worker_node_ips, infra_node_ips, vip1, vip2):
    data = {
        "Host-Description": ["LB IP", "LDAP url", "Node Info"],
        "Host-Ips": [f"VIP1 {vip1} \n VIP2 {vip2}", "ldap-url", f"Bastion Nodes: {', '.join(bastion_node_ips)} \n Bootstrap Nodes: {', '.join(bootstrap_node_ips)} \n Master Nodes: {', '.join(master_node_ips)} \n Worker Nodes: {', '.join(worker_node_ips)} \n Infra Nodes: {', '.join(infra_node_ips)}"],
        "Details": ["Vip1 and Vip2", "LDAP url", "Node Ip info"]
    }

    df = pd.DataFrame(data)
    with pd.ExcelWriter("far.xlsx") as writer:
        df.to_excel(writer, sheet_name="Description", index=False)
        