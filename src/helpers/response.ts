import { Request, Response } from "express";
import jwt from "jsonwebtoken";
export const handleSuccess = (
  res: Response,
  data: Object,
  message?: string
) => {
  return res.send({
    code: 1,
    message: message,
    data: data,
  });
};

export const handleError = (
  res: Response,
  message: Object,
  status?: number
) => {
  if (status) {
    res.status(status);
  }
  return res.send({
    code: 2,
    message: message,
  });
};

export const errorSystem = (req: Request, res: Response, err: any) => {
  if (res)
    res.send({
      status: 0,
      code: 0,
      message: err.message,
    });
};
export const getCurrentId = async (req) => {
  return new Promise((resolve, reject) => {
    var id = "";
    var token =
      req.body.token ||
      req.query.token ||
      req.headers["x-access-token"] ||
      req.cookies.token;
    if (token && token.search("token=") !== -1) {
      token = token.substring(token.search("token=") + 6);
    }
    if (!token) {
      id = "";
    } else {
      jwt.verify(token, "minigames", function (err, decoded) {
        if (err) {
          id = "";
        } else {
          id = decoded.id;
        }
      });
    }
    resolve(id);
  });
};
