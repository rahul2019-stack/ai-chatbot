from fastapi import FastAPI

from helpers import create_excel_sheet
from models.input import userInputs

app = FastAPI()

@app.post("/createFar")
def create_far(user: userInputs):
    create_excel_sheet(user.env, user.app, user.department, user.bastion_node_ips, user.bootstrap_node_ips, user.master_node_ips, 
                       user.worker_node_ips, user.infra_node_ips, user.vip1, user.vip2, user.dns_name)
    return {
        "message": f"Hello {user.env}, your bastion node IPs are {user.bastion_node_ips}, bootstrap node IPs are {user.bootstrap_node_ips}, master node IPs are {user.master_node_ips}, worker node IPs are {user.worker_node_ips}, infra node IPs are {user.infra_node_ips}, VIP1 is {user.vip1}, and VIP2 is {user.vip2}."
    }
