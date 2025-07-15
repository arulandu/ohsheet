from fastapi import APIRouter
from .index import router as index_router
from .relate import router as relate_router
from .query import router as query_router

router = APIRouter()

router.include_router(index_router)
router.include_router(relate_router)
router.include_router(query_router) 