/** Word add-in entry point for the shared Google sign-in (see src/shared/auth.ts). */
import { configureAuth } from "../shared/auth";

configureAuth("zeeai.word", "Word");

export * from "../shared/auth";
