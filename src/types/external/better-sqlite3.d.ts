declare module 'better-sqlite3' {
  export default class Database {
    constructor(path?: string);
    prepare(sql: string): Statement;
    exec(sql: string): this;
    pragma(pragma: string): any;
    transaction<T extends (...args: any[]) => any>(fn: T): T;
  }

  export class Statement {
    run(...args: any[]): any;
    get(...args: any[]): any;
    all(...args: any[]): any[];
  }
}
