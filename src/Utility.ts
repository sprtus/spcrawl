import path from 'path';

export class Utility {
  // Build path
  static path(filePath: string, cwd: boolean = true): string {
    return path.normalize(`${cwd ? process.cwd() : __dirname}/${filePath}`);
  }
};
