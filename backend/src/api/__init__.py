from fastapi import APIRouter
from .index import router as index_router
from .relate import router as relate_router

router = APIRouter()

router.include_router(index_router)
router.include_router(relate_router) 