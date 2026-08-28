/// <reference types="@rsbuild/core/types" />

declare module '*.css';
declare module '*.base64' {
  const value: string;
  export default value;
}
