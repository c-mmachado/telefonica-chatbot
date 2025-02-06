import { ConnectionPool } from "mssql";

export declare interface Technician {
  id: number;
  email: string;
  fecha: Date;
  activo: boolean;
}

export class TechnicianRepository {
  constructor(private readonly _db: ConnectionPool) {}

  public async technicians(): Promise<Technician[]> {
    if (!this._db.connected) {
      // If the connection is not open, open it
      await this._db.connect();
    }

    try {
      console.debug(
        `[${TechnicianRepository.name}][DEBUG] [${this.technicians.name}] Fetching technician emails...`
      );

      // Perform the query to get the technicians from the database
      const result = await this._db
        .query<Technician>`SELECT * FROM dbo.tecnicalmails WHERE activo = 1`;

      // Return the result
      if (result?.recordset) {
        return result.recordset.map((r: Technician) => {
          return {
            id: r.id,
            email: r.email,
            fecha: new Date(r.fecha),
            activo: Boolean(r.activo),
          };
        });
      }
      return [];
    } catch (error: any) {
      // Catches any errors that occur during the technicians query
      console.error(
        `[${TechnicianRepository.name}][ERROR] [${
          this.technicians.name
        }] error:\n${JSON.stringify(error, null, 2)}`
      );

      // Rethrows the error to the caller
      throw error;
    }
  }

  public async technician(id: number): Promise<Technician> {
    if (!this._db.connected) {
      // If the connection is not open, open it
      await this._db.connect();
    }

    try {
      console.debug(
        `[${TechnicianRepository.name}][DEBUG] [${this.technicians.name}] Fetching technician emails...`
      );

      // Perform the query to get the technician from the database
      const result = await this._db
        .query<Technician>`SELECT * FROM dbo.tecnicalmails WHERE activo = 1 AND id = ${id}`;

      // Return the result
      if (result?.recordset) {
        return result.recordset.map((r: Technician) => {
          return {
            id: r.id,
            email: r.email,
            fecha: new Date(r.fecha),
            activo: Boolean(r.activo),
          };
        })[0];
      }
      return {} as Technician;
    } catch (error: any) {
      // Catches any errors that occur during the technicians query
      console.error(
        `[${TechnicianRepository.name}][ERROR] [${
          this.technicians.name
        }] error:\n${JSON.stringify(error, null, 2)}`
      );

      // Rethrows the error to the caller
      throw error;
    }
  }

  public async createTechnician(email: string): Promise<any> {
    if (!this._db.connected) {
      // If the connection is not open, open it
      await this._db.connect();
    }

    try {
      console.debug(
        `[${TechnicianRepository.name}][DEBUG] [${this.createTechnician.name}] email: ${email}`
      );

      // Perform the query to create the technician in the database
      const result = await this._db
        .query<any>`INSERT INTO dbo.tecnicalmails (email, activo) VALUES (${email}, 1)`;

      return result;
    } catch (error: any) {
      // Catches any errors that occur during the technician creation query

      console.error(
        `[${TechnicianRepository.name}][ERROR] [${
          this.createTechnician.name
        }] error:\n${JSON.stringify(error, null, 2)}`
      );

      // Rethrows the error to the caller
      throw error;
    }
  }

  public async updateTechnician(body: {
    id: number;
    email: string;
    activo: number;
  }): Promise<any> {
    if (!this._db.connected) {
      // If the connection is not open, open it
      await this._db.connect();
    }

    try {
      console.debug(
        `[${TechnicianRepository.name}][DEBUG] [${this.updateTechnician.name}] id: ${body.id}`
      );

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
      const result = await this._db.query<any>(
        `${query} WHERE id = ${body.id}`
      );

      return result;
    } catch (error: any) {
      // Catches any errors that occur during the technician update query

      console.error(
        `[${TechnicianRepository.name}][ERROR] [${
          this.updateTechnician.name
        }] error:\n${JSON.stringify(error, null, 2)}`
      );

      // Rethrows the error to the caller
      throw error;
    }
  }

  public async deleteTechnician(id: number): Promise<any> {
    if (!this._db.connected) {
      // If the connection is not open, open it
      await this._db.connect();
    }

    try {
      console.debug(
        `[${TechnicianRepository.name}][DEBUG] [${this.deleteTechnician.name}] id: ${id}`
      );

      // Perform the query to delete the technician in the database
      const result = await this._db
        .query<any>`DELETE FROM dbo.tecnicalmails WHERE id = ${id}`;

      return result;
    } catch (error: any) {
      // Catches any errors that occur during the technician deletion query

      console.error(
        `[${TechnicianRepository.name}][ERROR] [${
          this.deleteTechnician.name
        }] error:\n${JSON.stringify(error, null, 2)}`
      );

      // Rethrows the error to the caller
      throw error;
    }
  }
}
