# SlideFormatter

SlideFormatter is a .NET Framework application designed to format and modify PowerPoint slides programmatically. The project leverages the DocumentFormat.OpenXml library to interact with PowerPoint presentation slides and provides various utilities to manipulate and enhance the content of the slides.

## Features

- **Update Slide Title**: Allows updating the title of a provided slide.
  
- **Transfer Text**: Transfers text content from text boxes to corresponding shapes within a slide.
  
- **Align & Resize Shapes**: Aligns and resizes shapes based on a reference shape's properties.
  
- **Font Modification**: Changes the font of all text boxes on a slide to "Beirut".
  
- **Textbox Alignment & Resizing**: Aligns and resizes all text boxes on a slide based on properties of a reference textbox.
  
- **Bullet Style Modification**: Changes the bullet style of paragraphs to dots and provides utilities to check bullet styles.
  
- **Styling Removal**: Provides functionality to remove bold and underline styles from text.
  
- **Error Logging**: Logs errors to a designated log file for debugging and tracking purposes.
  
- **GUI Interaction**: A button click event handler to allow users to process a chosen PowerPoint slide from the application interface.

## Prerequisites

- .NET Framework
- DocumentFormat.OpenXml library

## Usage

1. Open the application.
2. Use the provided GUI to select and process PowerPoint slides.
3. Monitor the application's feedback and error logs for insights into processing results and potential issues.

## Contributing

Contributions are welcome! Please fork the repository and create a pull request with your changes.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.

