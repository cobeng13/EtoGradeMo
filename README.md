# EtoGradeMo

Upload the macro-enabled class record and enter the report date, section, course code, and course title. Download one Word file containing the selected template pages for every named student. Full Grade Breakdown places each student's Midterm and Finals pages together.

## Run with Docker

    docker compose up -d --build --wait

Open http://localhost:8006. The service binds to localhost by default. For online use, put it behind the same authenticated HTTPS proxy as Student Progress Maker and set the request limit to 11 MB. Set ETOGRADE_PORT or ETOGRADE_BIND_ADDRESS in a local .env file to change the published port or address.

On the deployment host, run `bash deploy/update.sh` from a clean Git checkout to pull, rebuild, and wait for the app to respond.

## Run with Python

    python -m pip install -r requirements.txt
    python -m uvicorn app:app --port 8006

## Workbook and template

Download the blank standardized class record from the top of the app and fill it in. Use the original headers in row 1. Recalculate and save it in Excel before upload. The app reads saved results without running macros. It uses the template's seven quiz and activity slots and reads the aggregate values saved in the workbook. Uploaded records are handled per request and excluded from the Docker build.

The optional Finals table calculates the exam score needed as (75 minus the weighted midterm and finals class performance) divided by 0.40. It uses the saved weighted grades and rounds up to two decimal places. A result above 100 means the exam alone cannot bring the grade to 75.

