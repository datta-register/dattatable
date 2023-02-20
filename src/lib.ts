// Export the library
export * from "./common";
export * from "./auditLog";
export * from "./dashboard";

// Export the library constant
import * as Common from "./common";
import * as AuditLog from "./auditLog";
import * as Dashboard from "./dashboard";
export const DattaTable = { ...Common, ...AuditLog, ...Dashboard }