# BCM Transformer

A web application for generating Business Capability Maps from Excel data and creating professional PowerPoint presentations.

## Features

- **Web Interface**: User-friendly web GUI for uploading Excel files and configuring presentation settings
- **Excel Integration**: Import business capability data from Excel spreadsheets
- **PowerPoint Generation**: Automatically create professional PowerPoint presentations
- **Customizable Styling**: Configure fonts, colors, borders, and dimensions
- **Command Line Interface**: Generate presentations directly from the command line
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

### Command Line Interface

Generate a PowerPoint presentation directly from the command line:

```bash
generate_presentation \
  --fontSizeLevel1 12 \
  --fontSizeLevel2 8 \
  --colorFillLevel1 "#023047" \
  --colorFillLevel2 "#FCBF49" \
  --textColorLevel1 "#FFFFFF" \
  --textColorLevel2 "#000000" \
  --borderColor "#000000" \
  --widthLevel2 2.7 \
  --heightLevel2 1.0 \
  --excelPath "path/to/your/data.xlsx" \
  --outputPath "output/presentation.pptx"
```

### Docker Deployment

Build and run with Docker:

```bash
docker build -t bcm_transformer .
docker run -p 5000:5000 bcm_transformer
```

## Excel Data Format

The application expects Excel files with business capability data. Use the provided example file as a template:

- Example file: `bcm_transformer/excel_data/bcm_test_source.xlsx`

## Configuration Options

| Parameter | Description | Default |
|-----------|-------------|---------|
| `fontSizeLevel1` | Font size for level 1 capabilities | 12 |
| `fontSizeLevel2` | Font size for level 2 capabilities | 8 |
| `colorFillLevel1` | Background color for level 1 | "#023047" |
| `colorFillLevel2` | Background color for level 2 | "#FCBF49" |
| `textColorLevel1` | Text color for level 1 | "#FFFFFF" |
| `textColorLevel2` | Text color for level 2 | "#000000" |
| `borderColor` | Border color for all elements | "#000000" |
| `widthLevel2` | Width of level 2 boxes | 2.7 |
| `heightLevel2` | Height of level 2 boxes | 1.0 |

## License

This project is licensed under the Polyform Noncommercial License 1.0.0. See the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For issues and questions, please open an issue on GitHub.
