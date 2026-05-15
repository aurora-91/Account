"""
main.py — CLI runner (batch / cron / thử nghiệm không cần web).

Ví dụ:
  python main.py                                   # chạy toàn pipeline
  python main.py --buoc kho
  python main.py --buoc so_ban --test 5
  python main.py --buoc doi_soat
  python main.py --kho /data/kho --bb /data/bb --out /data/out
"""
import argparse
import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from config import Config
from modules import doi_soat, kho, so_ban

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(message)s',
    datefmt='%H:%M:%S',
)

_STEPS = {'kho': kho.run, 'so_ban': so_ban.run, 'doi_soat': doi_soat.run}


def _run_step(name: str, config: Config) -> bool:
    r = _STEPS[name](config)
    if r['success']:
        logging.info(f"✓ {r['message']}")
        if r.get('output_file'):
            logging.info(f"  → {r['output_file']}")
    else:
        logging.error(f"✗ {r['message']}")
        for e in r.get('errors', []):
            logging.error(f"  - {e}")
    return r['success']


def main():
    p = argparse.ArgumentParser(description='SIM Tool — CLI')
    p.add_argument('--buoc', choices=['kho', 'so_ban', 'doi_soat', 'all'], default='all')
    p.add_argument('--kho',  metavar='DIR')
    p.add_argument('--bb',   metavar='DIR')
    p.add_argument('--out',  metavar='DIR')
    p.add_argument('--test', type=int, default=0, metavar='N')
    args = p.parse_args()

    cfg = Config()
    if args.kho:  cfg.thu_muc_kho     = args.kho
    if args.bb:   cfg.thu_muc_bienban = args.bb
    if args.out:  cfg.output_folder   = args.out
    if args.test: cfg.so_ban_so_file_test = args.test

    steps = ['kho', 'so_ban', 'doi_soat'] if args.buoc == 'all' else [args.buoc]
    ok = all(_run_step(s, cfg) for s in steps)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
