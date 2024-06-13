import excelRouter from "./excelRouter";

export const router = (app: any) => {
  app.use("/api/excel", excelRouter);
};
