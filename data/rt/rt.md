# Request Tracker (RT) REST 2.0 API - OpenAPI Specification Summary

## Overview

This is an **OpenAPI 3.0.3** specification for the **Request Tracker (RT) REST 2.0 API**, reverse-engineered from the official Best Practical RT 5.0.4 REST 2 documentation at https://docs.bestpractical.com/rt/5.0.4/RT/REST2.html

**Server (defined in spec)**  
`https://pre-epg-vmticket-01.hi.inet` _(Development RT server – replace with your production instance in actual use)_

All requests **must use HTTPS**.

## Authentication Methods

The API supports three authentication mechanisms (all globally applied):

1. **Token Auth**  
   Header: `Authorization: token <your_token>`  
   or query: `?token=<your_token>`

2. **Basic Auth**  
   Standard HTTP Basic (RT username + password)

3. **Cookie Auth**  
   Cookie name: `RT_SID` (re-use from web login)

## Supported Resources / Tags

| Tag            | Description                                                    | Extensions Required |
| -------------- | -------------------------------------------------------------- | ------------------- |
| General        | Root and system info                                           | —                   |
| Queues         | Queues management                                              | —                   |
| Tickets        | Full ticket CRUD + correspond/comment/bulk/history/attachments | —                   |
| Users          | User management                                                | —                   |
| Groups         | Group management                                               | —                   |
| Custom Fields  | Custom fields                                                  | —                   |
| Custom Roles   | Custom roles                                                   | —                   |
| Assets         | Asset tracking                                                 | Assets extension    |
| Catalogs       | Asset catalogs                                                 | Assets extension    |
| Articles       | Knowledge-base articles                                        | Articles extension  |
| Classes        | Article classes                                                | Articles extension  |
| Transactions   | History / audit log entries                                    | —                   |
| Attachments    | File attachments                                               | —                   |
| Saved Searches | Saved search management                                        | —                   |

## Endpoints

### General
| Method | Endpoint   | Summary                  | Description                                      |
|--------|------------|--------------------------|--------------------------------------------------|
| GET    | /          | API Root                 | Returns hyperlinks to available resources        |
| GET    | /rt        | System Information       | Returns RT system information                    |

### Queues
| Method | Endpoint              | Summary                | Description                                      |
|--------|-----------------------|------------------------|--------------------------------------------------|
| GET    | /queues/all           | List All Queues        | Include disabled queues with ?find_disabled_rows=true |
| GET    | /queues               | Search Queues          | Search with query, orderby, pagination           |
| POST   | /queues               | Create Queue           | Create a new queue                               |
| GET    | /queue/{id}           | Get Queue by ID        | Queue details                                    |
| PUT    | /queue/{id}           | Update Queue           | Update queue by ID                               |
| DELETE | /queue/{id}           | Delete Queue           | Delete queue by ID                               |
| GET    | /queue/{name}         | Get Queue by Name      | Queue details by name                            |
| PUT    | /queue/{name}         | Update Queue by Name   | Update queue by name                             |

### Tickets
| Method | Endpoint                       | Summary                     | Description                                      |
|--------|----------------------------|---------------------------------|--------------------------------------------------|
| GET    | /tickets                       | Search Tickets              | TicketSQL query, format, orderby, pagination    |
| POST   | /tickets                       | Create Ticket               | Create a new ticket                              |
| GET    | /ticket/{id}                   | Get Ticket                  | Ticket details                                   |
| PUT    | /ticket/{id}                   | Update Ticket               | Update ticket                                    |
| DELETE | /ticket/{id}                   | Delete Ticket               | Delete ticket                                    |
| POST   | /ticket/{id}/correspond        | Add Correspondence to Ticket| Add public reply + attachments (multipart OK)    |
| POST   | /ticket/{id}/comment           | Add Comment to Ticket       | Add private comment + attachments (multipart OK)|
| GET    | /ticket/{id}/history           | Get Ticket History          | Transaction history (paginated)                  |
| GET    | /ticket/{id}/attachments      | Get Ticket Attachments      | List all attachments on ticket                   |
| PUT    | /tickets/bulk                  | Bulk Update Tickets         | Update multiple tickets in one request           |

### Users
| Method | Endpoint              | Summary                | Description                                      |
|--------|-----------------------|------------------------|--------------------------------------------------|
| GET    | /users                | Search Users           | Search with query, pagination                    |
| POST   | /users                | Create User            | Create a new user                                |
| GET    | /user/{id}            | Get User by ID         | User details                                     |
| PUT    | /user/{id}            | Update User            | Update user by ID                                |
| DELETE | /user/{id}            | Disable User           | Disable (not delete) user by ID                  |
| GET    | /user/{username}      | Get User by Username   | User details by username                         |

### Groups
| Method | Endpoint         | Summary       | Description                                      |
|--------|------------------|---------------|--------------------------------------------------|
| GET    | /groups          | Search Groups | Search with query, pagination                    |
| POST   | /groups          | Create Group  | Create a new group                               |
| GET    | /group/{id}      | Get Group     | Group details                                    |
| PUT    | /group/{id}      | Update Group  | Update group                                     |
| DELETE | /group/{id}      | Disable Group | Disable (not delete) group                       |

### Custom Fields
| Method | Endpoint               | Summary               | Description                                      |
|--------|------------------------|-----------------------|--------------------------------------------------|
| GET    | /customfields          | Search Custom Fields  | Search with query, pagination                    |
| POST   | /customfields          | Create Custom Field   | Create a new custom field                        |
| GET    | /customfield/{id}      | Get Custom Field      | Custom field details                             |
| PUT    | /customfield/{id}      | Update Custom Field   | Update custom field                              |
| DELETE | /customfield/{id}      | Disable Custom Field  | Disable custom field                             |

### Custom Roles
| Method | Endpoint              | Summary              | Description                                      |
|--------|-----------------------|----------------------|--------------------------------------------------|
| GET    | /customroles          | Search Custom Roles  | Search with query, pagination                    |
| POST   | /customroles          | Create Custom Role   | Create a new custom role                         |
| GET    | /customrole/{id}      | Get Custom Role      | Custom role details                              |
| PUT    | /customrole/{id}      | Update Custom Role   | Update custom role                               |
| DELETE | /customrole/{id}      | Disable Custom Role  | Disable custom role                               |

### Assets (requires Assets extension)
| Method | Endpoint             | Summary         | Description                                      |
|--------|----------------------|-----------------|--------------------------------------------------|
| GET    | /assets              | Search Assets   | Search with query, pagination                    |
| POST   | /assets              | Create Asset    | Create a new asset                              |
| GET    | /asset/{id}          | Get Asset       | Asset details                                    |
| PUT    | /asset/{id}          | Update Asset    | Update asset                                     |
| DELETE | /asset/{id}          | Delete Asset    | Delete asset                                     |
| GET    | /asset/{id}/history  | Get Asset History | Transaction history for asset                    |

### Catalogs (requires Assets extension)
| Method | Endpoint            | Summary             | Description                                      |
|--------|---------------------|---------------------|--------------------------------------------------|
| GET    | /catalogs           | Search Catalogs     | Search with query, pagination                    |
| POST   | /catalogs           | Create Catalog      | Create a new catalog                             |
| GET    | /catalog/{id}       | Get Catalog         | Catalog details by ID                            |
| PUT    | /catalog/{id}       | Update Catalog      | Update catalog                                   |
| DELETE | /catalog/{id}       | Disable Catalog     | Disable catalog                                  |
| GET    | /catalog/{name}     | Get Catalog by Name | Catalog details by name                          |

### Articles (requires Articles extension)
| Method | Endpoint              | Summary          | Description                                      |
|--------|-----------------------|------------------|--------------------------------------------------|
| GET    | /articles             | Search Articles  | Search with query, pagination                    |
| POST   | /articles             | Create Article   | Create a new article                             |
| GET    | /article/{id}         | Get Article      | Article details                                   |
| PUT    | /article/{id}         | Update Article   | Update article                                   |
| DELETE | /article/{id}         | Delete Article   | Delete article                                   |
| GET    | /article/{id}/history | Get Article History | Transaction history for article                 |

### Classes (requires Articles extension)
| Method | Endpoint           | Summary         | Description                                      |
|--------|----------------------|-----------------|--------------------------------------------------|
| GET    | /classes             | Search Classes  | Search with query, pagination                    |
| POST   | /classes             | Create Class    | Create a new class                               |
| GET    | /class/{id}          | Get Class       | Class details by ID                              |
| PUT    | /class/{id}          | Update Class    | Update class                                     |
| DELETE | /class/{id}          | Disable Class   | Disable class                                    |
| GET    | /class/{name}        | Get Class by Name | Class details by name                          |

### Transactions
| Method | Endpoint                        | Summary                    | Description                                      |
|--------|--------------------------------|----------------------------|--------------------------------------------------|
| GET    | /transactions                   | Search Transactions         | Search with query, pagination                     |
| GET    | /transaction/{id}              | Get Transaction            | Transaction details                               |
| GET    | /transaction/{id}/attachments  | Get Transaction Attachments  | List attachments on a transaction                  |

### Attachments
| Method | Endpoint          | Summary             | Description                                      |
|--------|-------------------|---------------------|--------------------------------------------------|
| GET    | /attachments       | Search Attachments   | Search with query, pagination                     |
| GET    | /attachment/{id}  | Get Attachment      | Attachment details (includes Base64 Content)      |

### Saved Searches
| Method | Endpoint        | Summary              | Description                                      |
|--------|----------------|----------------------|--------------------------------------------------|
| GET    | /searches      | List Saved Searches  | List all saved searches                          |
| POST   | /searches      | Create Saved Search  | Create a new saved search                        |
| GET    | /search/{id}   | Get Saved Search     | Saved search details                              |
| PUT    | /search/{id}   | Update Saved Search   | Update saved search                              |
| DELETE | /search/{id}   | Delete Saved Search   | Delete saved search                               |

## Notes from the Spec

-   This spec is **reverse-engineered** and may not cover every edge case – always cross-reference the official RT REST 2.0 documentation. This OpenAPI file provides a ready-to-use contract for generating clients (OpenAPI Generator, Swagger UI, Postman collections, etc.) to interact with RT's REST API.
-   Some endpoints may require specific RT extensions (e.g., Assets, Articles) to be installed on your RT instance.

You can visualize and interact with this OpenAPI specification using tools like [Swagger UI](https://editor.swagger.io/).
