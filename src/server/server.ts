import { readdirSync, lstatSync } from "fs";
import { join } from "path";
import { Server } from "http";
// import https, { ServerOptions } from "https";

import express, { Express, Router } from "express";

import { authMiddleware } from "./middleware/auth.middleware";
import { config } from "../config/config";
import helmet from "helmet";

// Create express application
export const app: Express = express();
app.use(express.json());
app.use(authMiddleware);
app.use(helmet());

async function discoverRoutes(router: Router, rootDir: string): Promise<void> {
    for (let fileName of readdirSync(rootDir)) {
        try {
            const filePath = join(rootDir, fileName);
            if (lstatSync(filePath).isDirectory()) {
                console.log(`Found route directory: ${fileName}, path: ${filePath}`);
                const dirRouter = Router();
                router.use(`/${fileName}`, dirRouter);
                await discoverRoutes(dirRouter, filePath);
            } else {
                if (fileName.endsWith(".ts") || fileName.endsWith(".js")) {
                    fileName = fileName.substring(0, fileName.indexOf("."));
                } else {
                    console.warn(`Skipping non .ts/.js file at path '${filePath}'`);
                    return;
                }

                console.log(`Found route file: ${fileName}, path: ${filePath}`);
                const module = await import(filePath);
                router.use(`/${fileName}`, module.router || module.default);
            }
        } catch (error: any) {
            console.error(`Error discovering routes for file at path ${fileName}:`, error);
            return Promise.reject(error);
        }
    }

    return Promise.resolve();
}

export async function startServer(routesDir?: string): Promise<void> {
    await discoverRoutes(app, routesDir || `${__dirname}/routes`).then(() => {
        console.log("Route discovery complete");
        // Create the server and listen on the specified port
        const server: Server =
            // https.createServer(options, app)
            app.listen(config.server.port, (): void => {
                console.info(`Bot started, '${app.name}' listening to`, server.address());
            });
    });
}

// Https server configuration
// const options: ServerOptions = {
//   key: config.ssl.key,
//   cert: config.ssl.cert,
// };

// Allow the auth-start.html and auth-end.html to be served from the public folder.
// expressApp.get(["/auth-start.html", "/auth-end.html"], async (req, res) => {
//   console.debug(`[expressApp][DEBUG] [${req.method}] req.url: ${req.url}`);
//   console.debug(
//     `[expressApp][DEBUG] [${req.method}] req.originalUrl:\n${JSON.stringify(req.originalUrl, null, 2)}`
//   );
//
//   send(
//     req,
//     path.join(
//       __dirname,
//       "public",
//       req.url.includes("auth-start.html") ? "auth-start.html" : "auth-end.html"
//     )
//   ).pipe(res);
// });
