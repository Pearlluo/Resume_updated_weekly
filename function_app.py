import logging
import importlib
import azure.functions as func

app = func.FunctionApp()

@app.schedule(
    schedule="0 0 18 * * SUN",
    arg_name="myTimer",
    run_on_startup=False,
    use_monitor=True
)
def weekly_resume_update(myTimer: func.TimerRequest) -> None:
    logging.info("Resume automation started")

    try:
        import GetResumeUpdated
        importlib.reload(GetResumeUpdated)

        logging.info("Resume automation completed")

    except Exception as e:
        logging.exception(f"Resume automation failed: {e}")
        raise