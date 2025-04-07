"""Command-line interface for the Spreadsheet MCP server."""

from .server import mcp


def main() -> None:
    """Run the Spreadsheet MCP server."""
    # parser = argparse.ArgumentParser(description="Spreadsheet MCP Server")
    # args = parser.parse_args(argv)

    print("Starting Spreadsheet MCP server")
    print("Press Ctrl+C to stop the server")

    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
