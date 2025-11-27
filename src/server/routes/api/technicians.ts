import { NextFunction, Router, Response, Request } from "express";

import { techRepository as repository } from "../../../config/db";

export const router = Router();

router.get("/", async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    try {
        // Fetch the technicians from the database
        const result = await repository.technicians();

        // Send the technicians as a JSON response
        res.status(200).json({
            status: 200,
            statusText: "OK",
            message: "Technicians fetched successfully",
            data: { technicians: result },
        });
    } catch (error: any) {
        console.error(error);

        const { name, ...rest } = error;
        if ("originalError" in rest) {
            delete rest.originalError;
        }

        // Send a 500 error to the client with the error
        res.status(500).json({
            status: 500,
            statusText: "Internal Server Error",
            message: "An error occurred while fetching technicians",
            data: {
                error: {
                    ...rest,
                    type: error.name,
                    message: error.message,
                },
            },
        });
    }

    next();
});

router.get("/:id", async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    try {
        // Parse the id from the request params
        const id = parseInt(req.params?.id);
        if (isNaN(id)) {
            // If the request params does not contain a valid id, send a 400 error to the client
            res.status(400).json({
                status: 400,
                statusText: "Bad Request",
                message: "Request parameter 'id' is required is required in the request path [/:id]",
                data: {
                    error: {
                        type: "InvalidRequestParameter",
                        message: "Request parameter 'id' is required in the request path [/:id]",
                        fields: {
                            id: {
                                in: "path",
                                required: true,
                                type: "number",
                            },
                        },
                    },
                },
            });
            return;
        }

        // Fetch the technicians from the database
        const result = await repository.technician(id);

        // Send the technicians as a JSON response
        res.status(200).json({
            status: 200,
            statusText: "OK",
            message: "Technician fetched successfully",
            data: { technician: result },
        });
    } catch (error: any) {
        console.error(error);

        const { name, ...rest } = error;
        if ("originalError" in rest) {
            delete rest.originalError;
        }

        // Send a 500 error to the client with the error
        res.status(500).json({
            status: 500,
            statusText: "Internal Server Error",
            message: "An error occurred while fetching technician",
            data: {
                error: {
                    ...rest,
                    type: error.name,
                    message: error.message,
                },
            },
        });
    }

    next();
});

router.post("/", async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    try {
        if (typeof req?.body?.email !== "string") {
            // If the request body does not contain a email field, send a 400 error to the client
            res.status(400).json({
                status: 400,
                statusText: "Bad Request",
                message: "Field 'email' is required in the request body",
                data: {
                    error: {
                        type: "InvalidRequestBody",
                        message: "Field 'email' is required in the request body",
                        fields: {
                            email: {
                                in: "body",
                                required: true,
                                type: "string",
                            },
                        },
                    },
                },
            });
            return;
        }

        console.debug("req.body:", req.body);

        // Creates a new technician in the database
        const result = await repository.createTechnician(req.body.email);

        // Send the technician creation result as a JSON response
        res.status(201).json({
            status: 201,
            statusText: "Created",
            message: "Technician created successfully",
            data: { technician: result },
        });
    } catch (error: any) {
        console.error(error);

        const { name, ...rest } = error;
        if ("originalError" in rest) {
            delete rest.originalError;
        }

        // Send a 500 error to the client with the error
        res.status(500).json({
            status: 500,
            statusText: "Internal Server Error",
            message: "An error occurred while creating technician",
            data: {
                error: {
                    ...rest,
                    type: error.name,
                    message: error.message,
                },
            },
        });
    }

    next();
});

router.put("/:id", async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    try {
        // Parse the id from the request params
        const id = parseInt(req.params.id);
        if (isNaN(id)) {
            // If the request params does not contain a valid id, send a 400 error to the client
            res.status(400).json({
                status: 400,
                statusText: "Bad Request",
                message: "Request parameter 'id' is required is required in the request path [/:id]",
                data: {
                    error: {
                        type: "InvalidRequestParameter",
                        message: "Request parameter 'id' is required in the request path [/:id]",
                        fields: {
                            id: {
                                in: "path",
                                required: true,
                                type: "number",
                            },
                        },
                    },
                },
            });
            return;
        }

        if (typeof req?.body !== "object") {
            // If the request does not contain a body, send a 400 error to the client
            res.status(400).json({
                status: 400,
                statusText: "Bad Request",
                message: "Missing request body",
                data: {
                    error: {
                        type: "MissingRequestBody",
                        message: "Request body is required",
                    },
                },
            });
            return;
        }

        console.debug("req.body:", req.body);

        // Updates a new technician in the database
        const result = await repository.updateTechnician({
            id: id,
            email: req.body.email,
            activo: req.body.activo,
        });

        // Send the technician update result as a JSON response
        res.status(200).json({
            status: 200,
            statusText: "OK",
            message: "Technician updated successfully",
            data: { technician: result },
        });
    } catch (error: any) {
        console.error(error);

        const { name, ...rest } = error;
        if ("originalError" in rest) {
            delete rest.originalError;
        }

        // Send a 500 error to the client with the error
        res.status(500).json({
            status: 500,
            statusText: "Internal Server Error",
            message: "An error occurred while updating technician",
            data: {
                error: {
                    ...rest,
                    type: error.name,
                    message: error.message,
                },
            },
        });
    }

    next();
});

router.delete("/:id", async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    console.debug(`req.headers:`, req.headers);

    try {
        // Parse the id from the request params
        const id = parseInt(req.params.id);
        if (isNaN(id)) {
            // If the request params does not contain a valid id, send a 400 error to the client
            res.status(400).json({
                status: 400,
                statusText: "Bad Request",
                message: "Request parameter 'id' is required is required in the request path [/:id]",
                data: {
                    error: {
                        type: "InvalidRequestParameter",
                        message: "Request parameter 'id' is required in the request path [/:id]",
                        fields: {
                            id: {
                                in: "path",
                                required: true,
                                type: "number",
                            },
                        },
                    },
                },
            });
            return;
        }

        if (typeof req?.body !== "object") {
            // If the request does not contain a body, send a 400 error to the client
            res.status(400).json({
                status: 400,
                statusText: "Bad Request",
                message: "Missing request body",
                data: {
                    error: {
                        type: "MissingRequestBody",
                        message: "Request body is required",
                    },
                },
            });
            return;
        }

        // Deletes a technician in the database
        const result = await repository.deleteTechnician(id ?? -1);

        // Send the technician deletion result as a JSON response
        res.status(200).json({
            status: 200,
            statusText: "OK",
            message: "Technician deleted successfully",
            data: { technician: result },
        });
    } catch (error: any) {
        console.error(error);

        const { name, ...rest } = error;
        if ("originalError" in rest) {
            delete rest.originalError;
        }

        // Send a 500 error to the client with the error
        res.status(500).json({
            status: 500,
            statusText: "Internal Server Error",
            message: "An error occurred while deleting technician",
            data: {
                error: {
                    ...rest,
                    type: error.name,
                    message: error.message,
                },
            },
        });
    }

    next();
});
