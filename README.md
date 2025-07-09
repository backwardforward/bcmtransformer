# BCM Transformer

> **Notice:** This project is **source-available** but **not open source**. The source code is provided for review and non-commercial use only, under the Polyform Noncommercial License 1.0.0. Commercial use, modification, or redistribution is not permitted.

A web application for generating Business Capability Maps from Excel data and creating professional PowerPoint presentations.

## Features

- **Web Interface**: User-friendly web GUI for uploading Excel files and configuring presentation settings
- **Excel Integration**: Import business capability data from Excel spreadsheets
- **PowerPoint Generation**: Automatically create professional PowerPoint presentations
- **Customizable Styling**: Configure fonts, colors, borders, and dimensions
- **Docker Support**: Easy deployment with Docker containers

## Installation

### Prerequisites

- Python 3.7 or higher
- pip package manager

### Setup

1. Clone the repository:
```bash
git clone https://github.com/backwardforward/bcmtransformer.git
cd bcmtransformer
```

2. Create and activate a virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install the package:
```bash
pip install -e .
```

## Usage

### Web Application

Start the web application:
```bash
bcm_transformer_app
# or
python -m bcm_transformer.app
```

Access the web interface at: [http://localhost:5000/](http://localhost:5000/)

### Docker Deployment

Build and run with Docker:

```bash
docker build -t bcm_transformer .
docker run -p 5000:5000 bcm_transformer
```

## Excel Data Format

The application expects Excel files with business capability data. Use the provided example file as a template:

- Example file: `bcm_transformer/excel_data/bcm_test_source.xlsx`

## License

- This project is **not open source**. The source code is provided under a source-available license.
- Licensed under the Polyform Noncommercial License 1.0.0. See the [LICENSE](LICENSE) file for details.
- **Commercial use is strictly prohibited.**

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For issues and questions, please open an issue on GitHub.
