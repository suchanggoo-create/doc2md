import logging


def configure_logging(level: str) -> None:
    level_map = {"error": logging.ERROR, "warn": logging.WARNING, "info": logging.INFO, "debug": logging.DEBUG}
    logging.basicConfig(
        level=level_map.get(level, logging.INFO),
        format="%(levelname)s %(message)s",
    )
