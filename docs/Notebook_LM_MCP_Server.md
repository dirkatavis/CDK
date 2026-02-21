1. Core MCP Server Architecture
An MCP server allows AI agents to interact with external tools and data through three primary capabilities:
• Tools (Primary Focus): Functions that the LLM can call to perform actions (e.g., creating a notebook or querying content).
• Resources: Read-only data like file contents or API responses.
• Prompts: Templates that help users accomplish specific tasks.
Important Technical Constraint: For STDIO-based servers, the server must never write to stdout (e.g., using print() or console.log()), as this corrupts JSON-RPC messages. All logging must be directed to stderr.

--------------------------------------------------------------------------------
2. NotebookLM API Endpoints (Enterprise)
The following methods are available via the NotebookLM Enterprise API for programmatic management. Your Gemini partner should implement these as Tools within the MCP server.
Notebook Management
• notebooks.create: Creates a new notebook with a specified title.
• notebooks.get: Retrieves details for a specific notebook ID.
• notebooks.listRecentlyViewed: Lists up to the last 500 accessed notebooks.
• notebooks.batchDelete: Deletes a notebook using its complete resource name.
• notebooks.share: Grants user roles (Owner, Writer, Reader) to specific email addresses.
Source Management
• notebooks.sources.batchCreate: Adds data sources in bulk. Supported types include Google Docs, Slides, raw text, web URLs, and YouTube videos.
• notebooks.sources.uploadFile: Uploads single files. Supported formats include .pdf, .txt, .md, .docx, .pptx, .xlsx, and various audio/image types.
• notebooks.sources.get: Retrieves a specific source using its SOURCE_ID.
• notebooks.sources.batchDelete: Deletes data sources in bulk.
Studio Artifacts
• Audio Overviews: Programmatically generate deep-dive podcasts.
• Other Artifacts: Methods exist for generating video explainers, briefing docs, flashcards, infographics, and slide decks.

--------------------------------------------------------------------------------
3. Authentication Framework
Authentication is critical for the server to access Google Cloud resources.
• User Identity (Development): Use Application Default Credentials (ADC) by running gcloud auth application-default login to create a local credential file.
• Service Account (Production): The preferred method for hosted applications. The service account requires specific IAM roles, such as the Cloud NotebookLM User role.
• OAuth 2.0: If the application needs to act on behalf of a user without sharing credentials, use an OAuth client ID and client secret.
• Browser Automation (Unofficial): Some existing community MCP servers use Playwright or Puppeteer to extract session cookies for personal/Pro accounts that lack a formal API.

--------------------------------------------------------------------------------
4. Implementation Roadmap for Gemini
Your partner can use several languages to build the server, with Python and TypeScript being the most common:
• Python (FastMCP): Use the FastMCP class to automatically generate tool definitions from Python type hints and docstrings.
• TypeScript: Use the MCP SDK with Node.js (v16+). Ensure a build script is defined in package.json to compile the code before the host (e.g., Claude for Desktop) connects.
• Testing: Use the MCP Inspector tool or connect it to Claude for Desktop by adding the server configuration to the claude_desktop_config.json file.
5. Existing Reference Implementations
Gemini should review these established patterns to avoid common pitfalls:
• Unified CLI/MCP: A Python-based project that combines a CLI (nlm) with an MCP server, supporting advanced features like deep research and automatic auth refreshes.
• Browser Automation MCP: A TypeScript implementation that uses Chrome automation to interface with the consumer version of NotebookLM.
• RPC Mapping: An unofficial Python client that maps internal RPC endpoints to bypass browser automation overhead.
Would you like me to create a tailored report summarizing these technical specifications, or perhaps a slide deck that Gemini can use to visualize the server architecture?