import { ConnectionPool } from "mssql";

export declare interface APILog {
  id: number;
  fecha: Date;
  txt: any;
}

export class LogsRepository {
  constructor(private readonly _db: ConnectionPool) {}

  public async logs(): Promise<APILog[]> {
    if (!this._db.connected) {
      // If the connection is not open, open it
      await this._db.connect();
    }

    try {
      console.debug(
        `[${LogsRepository.name}][DEBUG] [${this.logs.name}] Fetching api logs...`
      );

      // Perform the query to get the api logs from the database
      const result = await this._db.query<
        APILog[]
      >`SELECT * FROM dbo.logschatbot`;

      // Return the result
      if (result?.recordset) {
        return result.recordset.map((r: APILog) => {
          return {
            id: r.id,
            fecha: new Date(r.fecha),
            txt: JSON.parse(r.txt),
          };
        });
      }
      return [];
    } catch (error: any) {
      // Catches any errors that occur during the technicians query

      console.error(
        `[${LogsRepository.name}][ERROR] [${
          this.logs.name
        }] error:\n${JSON.stringify(error, null, 2)}`
      );

      // Rethrows the error to the caller
      throw error;
    }
  }

  public async createLog(message: string): Promise<any> {
    if (!this._db.connected) {
      // If the connection is not open, open it
      await this._db.connect();
    }

    try {
      console.debug(
        `[${LogsRepository.name}][DEBUG] [${this.createLog.name}] Creating log with message: ${message}`
      );

      // Perform the query to create the api log in the database
      const result = await this._db
        .query<any>`INSERT INTO dbo.logschatbot (txt) VALUES (${message})`;

      return result;
    } catch (error: any) {
      // Catches any errors that occur during the api log creation query

      console.error(
        `[${LogsRepository.name}][ERROR] [${
          this.logs.name
        }] error:\n${JSON.stringify(error, null, 2)}`
      );

      // Rethrows the error to the caller
      throw error;
    }
  }
}
