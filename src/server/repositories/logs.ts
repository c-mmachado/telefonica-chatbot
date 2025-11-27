import { ConnectionPool, IResult } from "mssql";

export interface APILog {
    id: number;
    fecha: Date;
    txt: any;
}

export class LogsRepository {
    constructor(private readonly _db: ConnectionPool) {}

    private async _connect(): Promise<void> {
        await this._db.connect();
    }

    public async logs(): Promise<APILog[]> {
        if (!this._db.connected) {
            // If the connection is not open, open it
            await this._connect();
        }

        console.debug(`Fetching API logs`);

        // Perform the query to get the API logs from the database
        return await this._db.query<APILog[]>`SELECT * FROM dbo.logschatbot`
            .then((result: IResult<APILog>): APILog[] => {
                // Parse the results of the query to an APILog array

                if (result?.recordset?.length > 0) {
                    // If the query returns results, map the results to the APILog type
                    return result.recordset.map((r: APILog) => {
                        return {
                            id: r.id,
                            fecha: new Date(r.fecha),
                            txt: JSON.parse(r.txt),
                        };
                    });
                }

                // If the query returns no results, return an empty array
                return [];
            })
            .catch((error: any) => {
                console.error(error);
                throw error;
            });
    }

    public async createLog(message: string): Promise<void> {
        if (!this._db.connected) {
            // If the connection is not open, open it
            await this._connect();
        }

        console.debug(`Creating API log with message: '${message}'`);

        // Perform the query to create the API log in the database
        await this._db.query<any>`INSERT INTO dbo.logschatbot (txt) VALUES (${message})`.catch((error: any) => {
            console.error(`Error creating API log:`, error);
            throw error;
        });
    }
}
