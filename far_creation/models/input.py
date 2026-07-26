from typing import List, Optional

from pydantic import BaseModel, Field


class userInputs(BaseModel):
    env: str = Field(default="dev/pre-prod/uat/prod")
    app: str
    department: str
    dns_name: str
    vcentre: str = Field(default="b1u03")
    bastion_node_ips: List[str]
    bootstrap_node_ips: List[str]
    master_node_ips: List[str]
    worker_node_ips: List[str]
    infra_node_ips: Optional[List[str]]
    vip1: str
    vip2: str
