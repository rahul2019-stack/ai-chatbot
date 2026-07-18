from pydantic import BaseModel
from typing import List

class MemoryExtractionResponse(BaseModel):
    memories: List[str]
