import { ConnectionPool } from "mssql";

import { config } from "../config/config";

// Create database connection pool
export const dbConnection: ConnectionPool = new ConnectionPool({
  server: config.dbHost,
  port: config.dbPort,
  user: config.dbUser,
  password: config.dbPassword,
  database: config.dbName,
  options: {
    encrypt: false,
    enableArithAbort: true,
  },
});
