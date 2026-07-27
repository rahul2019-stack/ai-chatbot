from fastapi import FastAPI

from helpers import create_excel_sheet
from models.input import userInputs

app = FastAPI()

@app.post("/createFar")
def create_far(user: userInputs):
    msg = create_excel_sheet(user.env, user.app, user.department, user.bastion_node_ips, user.bootstrap_node_ips, user.master_node_ips, 
                       user.worker_node_ips, user.infra_node_ips, user.vip1, user.vip2, user.dns_name, user.vcentre)
    return {
        "message": msg
    }
