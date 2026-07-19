from typing import List

from pydantic import BaseModel


class userInputs(BaseModel):
    env: str
    app: str
    department: str
    bastion_node_ips: List[str]
    bootstrap_node_ips: List[str]
    master_node_ips: List[str]
    worker_node_ips: List[str]
    infra_node_ips: List[str]
    vip1: str
    vip2: str