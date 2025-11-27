import { ConnectionPool, IResult } from "mssql";

export declare interface Technician {
    id: number;
    email: string;
    fecha: Date;
    activo: boolean;
}

export class TechnicianRepository {
    constructor(private readonly _db: ConnectionPool) {}

    private async _connect(): Promise<void> {
        await this._db.connect();
    }

    public async technicians(): Promise<Technician[]> {
        if (!this._db.connected) {
            // If the connection is not open, open it
            await this._connect();
        }

        console.debug(`Fetching technicians`);

        // Perform the query to get the technicians from the database
        const result = await this._db.query<Technician>`SELECT * FROM dbo.tecnicalmails WHERE activo = 1`
            .then((result: IResult<Technician>): Technician[] => {
                console.debug(`result:`, result);

                if (result?.recordset?.length > 0) {
                    // Parse the recordset into a Technician array
                    return result.recordset.map((r: Technician) => {
                        return {
                            id: r.id,
                            email: r.email,
                            fecha: new Date(r.fecha),
                            activo: Boolean(r.activo),
                        };
                    });
                }

                // Return an empty array if no technicians are found
                return [];
            })
            .catch((error: any): never => {
                console.error(error);
                throw error;
            });

        // Return the result
        return result || [];
    }

    public async technician(id: number): Promise<Technician | null> {
        if (!this._db.connected) {
            // If the connection is not open, open it
            await this._connect();
        }

        console.debug(`Fetching technician by id:`, id);

        // Perform the query to get the technician from the database
        const result: Technician | null = await this._db
            .query<Technician>`SELECT * FROM dbo.tecnicalmails WHERE activo = 1 AND id = ${id}`
            .then((result: IResult<Technician>): Technician | null => {
                console.debug(`result:`, result);

                if (result?.recordset?.length > 0) {
                    // Parse the recordset into a Technician object
                    return result.recordset.map((r: Technician) => {
                        return {
                            id: r.id,
                            email: r.email,
                            fecha: new Date(r.fecha),
                            activo: Boolean(r.activo),
                        };
                    })[0];
                }

                // Return null if no technician is found
                return null;
            })
            .catch((error: any): never => {
                console.error(error);
                throw error;
            });

        // Return the result
        return result;
    }

    public async technicianByEmail(email: string): Promise<Technician | null> {
        if (!this._db.connected) {
            // If the connection is not open, open it
            await this._connect();
        }

        console.debug(`Fetching technician by email:`, email);

        // Perform the query to get the technician from the database
        const result: Technician | null = await this._db
            .query<Technician>`SELECT * FROM dbo.tecnicalmails WHERE activo = 1 AND email = ${email}`
            .then((result: IResult<Technician>): Technician | null => {
                console.debug(`result:`, result);

                if (result?.recordset?.length > 0) {
                    // Parse the recordset into a Technician object
                    return result.recordset.map((r: Technician) => {
                        return {
                            id: r.id,
                            email: r.email,
                            fecha: new Date(r.fecha),
                            activo: Boolean(r.activo),
                        };
                    })[0];
                }

                // Return null if no technician is found
                return null;
            })
            .catch((error: any): never => {
                console.error(error);
                throw error;
            });

        // Return the result
        return result;
    }

    public async createTechnician(email: string): Promise<void> {
        if (!this._db.connected) {
            // If the connection is not open, open it
            await this._connect();
        }

        // Perform the query to create the technician in the database
        await this._db.query<any>`INSERT INTO dbo.tecnicalmails (email, activo) VALUES (${email}, 1)`.catch(
            (error: any): never => {
                console.error(error);
                throw error;
            }
        );
    }

    public async updateTechnician(body: Partial<Technician>): Promise<void> {
        if (!this._db.connected) {
            // If the connection is not open, open it
            await this._connect();
        }

        console.debug(`id: `, body.id, ` email: `, body.email, ` activo: `, body.activo);

        if (!body.id) {
            throw new Error(`Field 'id' is required to update a technician`);
        }

        let query = `UPDATE dbo.tecnicalmails SET`;
        if (body.email) {
            query += ` email = ${body.email}`;
        }
        if (body.email && body.activo) {
            query += `, activo = ${body.activo}`;
        } else if (body.activo) {
            query += ` activo = ${body.activo}`;
        }

        // Perform the query to update the technician in the database
        await this._db.query<any>(`${query} WHERE id = ${body.id}`).catch((error: any): never => {
            console.error(error);
            throw error;
        });
    }

    public async deleteTechnician(id: number): Promise<void> {
        if (!this._db.connected) {
            // If the connection is not open, open it
            await this._connect();
        }

        console.debug(`id: `, id);

        // Perform the query to delete the technician in the database
        await this._db.query<any>`DELETE FROM dbo.tecnicalmails WHERE id = ${id}`.catch((error: any): never => {
            console.error(error);
            throw error;
        });
    }
}
