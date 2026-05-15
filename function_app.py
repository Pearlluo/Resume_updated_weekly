import logging
import azure.functions as func

import GetResumeUpdated


app = func.FunctionApp()


# Monday 2AM Perth Time
# Perth UTC+8 = Sunday 18:00 UTC
@app.schedule(
    schedule="0 0 18 * * SUN",
    arg_name="myTimer",
    run_on_startup=False,
    use_monitor=True
)
def weekly_resume_update(myTimer: func.TimerRequest) -> None:

    logging.info("====================================")
    logging.info("Resume automation started")
    logging.info("Perth Monday 2AM")
    logging.info("====================================")

    try:
        import importlib

        # 强制重新执行主程序
        importlib.reload(GetResumeUpdated)

        logging.info("Resume automation completed")

    except Exception as e:
        logging.exception(f"Resume automation failed: {e}")
        raise