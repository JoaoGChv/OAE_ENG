"""Command line entry point for delivery manager."""
from oae.ui import main
import logging

if __name__ == "__main__":
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s %(levelname)s [%(name)s] %(message)s"
    )
    main() 