// Initializes the logging setup for the application
import { logger } from "./utils/logging";

import "isomorphic-fetch";

import { startServer } from "./server/server";

logger.info("Starting Ticket Bot server...");
startServer();
