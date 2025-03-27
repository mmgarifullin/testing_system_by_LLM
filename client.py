import os
import json
from http.client import HTTPConnection
from dataclasses import dataclass, asdict

from typing import Dict, Union, Tuple, List


@dataclass
class Config:
    port: int = 5000
    host: str = "127.0.0.1"
    divider: float = 2.
    use_mult: bool = False
    atten_coeff: float = 1.
    dense_coeff: float = 1.
    sparse_coeff: float = 1.
    colbert_coeff: float = 1.

    @classmethod
    def load(cls, path: str = "config.json"):
        exe_path = os.path.abspath(os.path.dirname(__file__))
        path = os.path.join(exe_path, path)
        if os.path.isfile(path):
            cfg = json.load(open(path, "r", encoding="utf-8"))
            cfg = cls(**cfg)
        else:
            print(f"{path} is not a file. Instantiate a default Config.")
            cfg = cls()
        return cfg

    def from_params(self, params: Union[None, Dict[str, str]]):
        if params is None:
            return self
        cfg = asdict(self)
        for k, v in params.items():
            if k in cfg:
                cfg[k] = v
        return Config(**cfg)

    def asdict(self):
        return asdict(self)


class Client(HTTPConnection):
    def __init__(self, host: str = "localhost",
                 port: Union[None, int] = None,
                 timeout: Union[None, float] = None,
                 source_address: Union[None, Tuple[str, int]] = None,
                 blocksize: int = 8192):
        super().__init__(host, port, timeout, source_address, blocksize)

    @classmethod
    def launch(cls):
        cfg = Config.load()
        return cls(host=cfg.host, port=cfg.port)

    def check(self):
        try:
            self.request(method="GET", url="")
            resp = self.getresponse()
            assert resp.status == 200
            return True
        except Exception as e:
            print(f"Exception:  {e}")
            return False

    def __call__(self, src: List[str], dst: List[str],
                 params: Union[None, Dict[str, str]] = None):
        obj = json.dumps({"src": src, "dst": dst, "params": params})
        header = {'Content-Type': 'application/json'}
        try:
            self.request(method="POST", url="",
                         body=obj.encode(), headers=header)
            resp = self.getresponse()
            if resp.status == 200:
                return True, json.loads(resp.readline().decode())
            return False, f"Response status: {resp.status}"
        except Exception as e:
            return False, f"Exception: {e}"


def main():
    pass


if __name__ == '__main__':
    main()
