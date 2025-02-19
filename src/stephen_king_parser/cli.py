import asyncio
import logging
from pathlib import Path
from typing import Optional

import click

from .core.parser import KingWorksParser
from .utils.config import load_config


@click.group()
@click.option("--config", type=click.Path(exists=True), help="Path to config file")
@click.option("--debug/--no-debug", default=False, help="Enable debug logging")
@click.pass_context
def cli(ctx: click.Context, config: Optional[str], debug: bool):
    """Stephen King Works Parser CLI"""
    ctx.ensure_object(dict)

    # Setup logging
    log_level = logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=log_level, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    # Load configuration
    ctx.obj["config"] = load_config(config)


@cli.command()
@click.option(
    "--output", type=click.Path(), default="output", help="Output directory for results"
)
@click.option(
    "--format",
    type=click.Choice(["csv", "html", "json"]),
    multiple=True,
    default=["csv", "html"],
    help="Output format(s)",
)
@click.pass_context
def parse(ctx: click.Context, output: str, format: tuple):
    """Parse Stephen King's works and save to specified format(s)"""
    config = ctx.obj["config"]
    parser = KingWorksParser(config)

    # Create output directory
    output_dir = Path(output)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Run parser
    asyncio.run(parser.run(output_dir, list(format)))


@cli.command()
@click.pass_context
def clear_cache(ctx: click.Context):
    """Clear the cache directory"""
    config = ctx.obj["config"]
    cache_dir = Path(config.cache_dir)
    if cache_dir.exists():
        for file in cache_dir.glob("*"):
            file.unlink()
        click.echo("Cache cleared successfully")
    else:
        click.echo("Cache directory does not exist")


if __name__ == "__main__":
    cli(obj={})
