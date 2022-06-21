from dataclasses import dataclass


@dataclass
class PLMAction:
    coordinates: list
    event: str
    time_sleep: float