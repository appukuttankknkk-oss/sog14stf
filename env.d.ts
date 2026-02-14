
// Augmenting the existing NodeJS.ProcessEnv to include API_KEY.
// This avoids the "Cannot redeclare block-scoped variable 'process'" error
// which occurs when 'process' is already defined by node types in the environment.
declare namespace NodeJS {
  interface ProcessEnv {
    API_KEY: string;
    [key: string]: string | undefined;
  }
}
