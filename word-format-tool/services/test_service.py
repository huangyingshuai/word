"""
模块名称：services.test_service
模块职责：识别准确率测试服务
"""
from core.number_grouper import calc_recognition_accuracy, STANDARD_TEST_CASES

def run_standard_test() -> tuple[float, list]:
    return calc_recognition_accuracy(STANDARD_TEST_CASES)