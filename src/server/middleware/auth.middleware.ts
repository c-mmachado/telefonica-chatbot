import { Response, Request, NextFunction } from "express";
import jwt, { JwtPayload } from "jsonwebtoken";

import { HttpHeaders } from "../../utils/http";

import { config } from "../../config/config";

function validateToken(
  token: string,
  secret: string
): Promise<JwtPayload | string> {
  return new Promise((resolve, reject: (reason?: any) => void) => {
    const payload = jwt.verify(token, secret);
    if (payload instanceof Error) {
      reject(payload);
    }
    resolve(payload);
  });
}

export async function authMiddleware(
  req: Request,
  res: Response,
  next: NextFunction
): Promise<void> {
  if (
    req.url.includes("/api/token") ||
    req.url.includes("/api/messages") ||
    req.url.includes("health")
  ) {
    // Skip authentication for the /token endpoint
    console.debug(`Skipping authentication for '${req.url}' endpoint`);
    next();
    return;
  }

  const authHeader: string | string[] =
    req.headers["authorization"] ||
    req.headers[HttpHeaders.Authorization] ||
    "";
  let token: string | undefined = undefined;
  if (Array.isArray(authHeader)) {
    // If the Authorization header is an array, log a warning and return 401
    console.warn(`Authorization header is an array:`, authHeader);
    token = authHeader[0]?.trim();
  } else if (typeof authHeader === "string") {
    // If the Authorization header is a string, split it to get the token
    token = authHeader.trim();
  } else {
    // If the Authorization header is neither a string nor an array, log a warning and return 401
    console.warn(
      `Authorization header is neither a string nor an array:`,
      authHeader
    );
    res.status(401).send({
      status: 401,
      data: { message: "Authorization header is not valid" },
    });
    return;
  }

  if (!token) {
    // If the token is not provided, log a warning and return 401
    console.warn(`No token provided in Authorization header`);
    res.status(401).send({
      status: 401,
      data: { message: "No token provided" },
    });
    return;
  }

  if (!token.toLowerCase().startsWith("bearer")) {
    // If the token does not start with "Bearer", log a warning and return 401
    console.warn(`Token does not start with "Bearer": '${token}'`);
    res.status(401).send({
      status: 401,
      data: { message: 'Token must be of type "Bearer"' },
    });
    return;
  }

  // Remove "Bearer" prefix from the token
  token = token.replace(/bearer /i, "");
  // .replace("bearer ", "")
  // .replace("BEARER ", "");

  try {
    // Verify the token using the secret from the config
    const decoded = await validateToken(token, config.server.jwt.secret);
    console.debug(`Token verified successfully:`, decoded);

    // Continue middleware chain execution
    next();
  } catch (error: any) {
    // If token verification fails, log the error
    console.error(error);

    if (error instanceof jwt.JsonWebTokenError) {
      // If the token verification fails, check if the error is a 'JsonWebTokenError'
      res.status(401).send({ status: 401, data: { message: error.message } });
      return;
    }

    // If the error is not a 'JsonWebTokenError', return a generic 401
    res.status(401).send({
      status: 401,
      data: { message: "Invalid token" },
    });
    return;
  }
}
