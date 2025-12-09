import fs from "fs";

import gradient from "gradient-string";
import "isomorphic-fetch";

async function start() {
    const banner = fs.readFileSync("./data/banner.txt", "utf8");
    const sunset = gradient(["#ff9966", "#ff5e62", "#ffa34e"]);
    console.log(sunset.multiline(banner));

    // Initializes the logging setup for the application
    const logger = await import("./utils/logging").then((mod) => mod.logger);
    logger.info("Starting Ticket Bot server...");

    const server = await import("./server/server").then((mod) => mod.startServer);
    await server();
}
start();
