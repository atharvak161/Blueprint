import { Router, type IRouter } from "express";
import healthRouter from "./health";
import projectsRouter from "./projects";
import phasesRouter from "./phases";
import tasksRouter from "./tasks";
import resourcesRouter from "./resources";
import holidaysRouter from "./holidays";

const router: IRouter = Router();

router.use(healthRouter);
router.use(projectsRouter);
router.use(phasesRouter);
router.use(tasksRouter);
router.use(resourcesRouter);
router.use(holidaysRouter);

export default router;
