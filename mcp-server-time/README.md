# mcp-server-time

MCP server for retrieving and calculating time-related information.

## Features

- **Get Current Time**: Retrieve current time in any timezone (defaults to local timezone)
- **Time Calculation**: Calculate dates and times based on natural language queries (e.g., "What day is next Thursday?", "What day of the week is February 1st next year?")

## Installation

To add this MCP server to your environment, add the following to your MCP config file:

```json
{
  "mcpServers": {
    "kzmshx.mcp-server-time": {
      "autoApprove": [],
      "args": ["/path/to/start.sh"],
      "command": "sh"
    }
  }
}
```

## Tools

### get_current_time

```ts
/**
 * Get the current time in a specified timezone
 *
 * @param {string} [timezone] - IANA timezone identifier (e.g., "Asia/Tokyo", "America/New_York")
 *                             If not provided, uses the local timezone
 * @returns {object} Current time information
 * @returns {string} .iso - ISO 8601 formatted timestamp
 * @returns {string} .timezone - The timezone used for the time calculation
 * @returns {number} .timestamp - Unix timestamp in milliseconds
 * @returns {object} .parts - Broken down time parts
 * @returns {number} .parts.year - Full year
 * @returns {number} .parts.month - Month (1-12)
 * @returns {number} .parts.day - Day of month (1-31)
 * @returns {number} .parts.hour - Hour (0-23)
 * @returns {number} .parts.minute - Minute (0-59)
 * @returns {number} .parts.second - Second (0-59)
 * @returns {number} .parts.millisecond - Millisecond (0-999)
 */
type GetCurrentTime = (timezone?: string) => {
  iso: string;
  timezone: string;
  timestamp: number;
  parts: {
    year: number;
    month: number;
    day: number;
    hour: number;
    minute: number;
    second: number;
    millisecond: number;
  };
};
```

### calculate_time

```ts
/**
 * Calculate time based on a natural language query using chrono parser
 *
 * @param {string} query - Natural language time query. Examples:
 *                        - "next thursday"
 *                        - "tomorrow at 3pm"
 *                        - "2024/04/01"
 *                        - "what day is february 1st next year"
 * @param {string} [timezone] - IANA timezone identifier for the calculation
 *                             If not provided, uses the local timezone
 * @returns {object} Time calculation result
 * @returns {string} .iso - ISO 8601 formatted timestamp
 * @returns {string} .timezone - The timezone used for the calculation
 * @returns {number} .timestamp - Unix timestamp in milliseconds
 * @returns {object} .parts - Broken down time parts
 * @returns {number} .parts.year - Full year
 * @returns {number} .parts.month - Month (1-12)
 * @returns {number} .parts.day - Day of month (1-31)
 * @returns {number} .parts.hour - Hour (0-23)
 * @returns {number} .parts.minute - Minute (0-59)
 * @returns {number} .parts.second - Second (0-59)
 * @returns {number} .parts.millisecond - Millisecond (0-999)
 */
type CalculateTime = (
  query: string,
  timezone?: string
) => {
  iso: string;
  timezone: string;
  timestamp: number;
  parts: {
    year: number;
    month: number;
    day: number;
    hour: number;
    minute: number;
    second: number;
    millisecond: number;
  };
};
```
