# User Permission Management System (RBAC Model)

## Core System Features
Role-Based Access Control (RBAC) system with predefined roles:
- 🛡️ **Super Admin**
- 🔑 **Admin**
- 👀 **Guest**
- 💰 **Loan Pricing**
- 💰 **User**

## Role Permission Matrix
### User Management Permissions
| Operation/Role       | Super Admin | Admin | Guest | Loan Pricing |
|----------------------|-------------|-------|-------|--------------|
| Create User          | ✔️          | ✔️    | ✖️    | ✖️           |
| Delete Regular User  | ✔️          | ✔️    | ✖️    | ✖️           |
| Delete Admin         | ✔️          | ✖️    | ✖️    | ✖️           |
| Modify User Role     | ✔️          | ✖️    | ✖️    | ✖️           |

### Page Access Permissions
Page Access Control

| Page Path        | Allowed Roles                  |
|------------------|--------------------------------|
| /chat            | Super Admin, Admin             |
| /user-management | Super Admin, Admin             |
| /loan-pricing    | Super Admin, Admin, Loan Pricing |
| /Test Creator    | All roles                      |
                 |
