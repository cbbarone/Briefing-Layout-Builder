import { Router, type IRouter } from "express";
import healthRouter from "./health.js";
import radarRouter from "./radar.js";

const router: IRouter = Router();

router.use(healthRouter);
router.use(radarRouter);

export default router;
