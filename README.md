# GPON/XGSPON OLT Autonomous Testing Framework

## Overview
This project is an automated provisioning and continuous learning testing framework for Nokia OLTs. It manages SFU ONT registration, stores inventory data, and executes test scenarios. The framework includes an evolution mechanism that identifies consistently successful tests and adapts them for stricter or more advanced future testing cycles.

## Features
* **Automated ONT Provisioning:** Simulates or executes SSH commands for Nokia ISAM OLTs to provision SFU ONTs.
* **Database Management:** Utilizes SQLite to track ONT inventory and test case execution history.
* **Continuous Learning:** Test scenarios that pass successfully are marked as "evolved" to trigger advanced validation logic in subsequent runs.

## Prerequisites
* Python 3.8 or higher
* Git

## Installation and Setup

1.  Clone the repository:
    ```bash
    git clone <your-repository-url>
    cd <your-repository-folder>
    ```

2.  Create and activate a virtual environment:
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows use: venv\Scripts\activate
    ```

3.  Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

## Usage
Run the main automation script to initialize the database, provision the ONT, and execute the test cycles:

```bash
python main.py
