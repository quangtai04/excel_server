import express from "express";
import * as excelController from "../controllers/excel";
const router = express.Router();

router
  .post("/createExcelVnedu", excelController.createExcelVnedu)
  .post("/sortData", excelController.sortData);

export default router;
