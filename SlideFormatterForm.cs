using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using System.IO;


namespace SlideFormatter
{
    public partial class SlideFormatterForm : Form
    {
        public SlideFormatterForm()
        {
            InitializeComponent();
        }

        private System.Windows.Forms.Label label;

        /// <summary>
        /// Updates the title of a given slide.
        /// </summary>
        /// <param name="slidePart">The slide part to be updated.</param>
        /// <exception cref="ArgumentNullException">Thrown when the slidePart is null.</exception>
        public void UpdateTitle(SlidePart slidePart)
        {
            if (slidePart == null)
                throw new ArgumentNullException(nameof(slidePart), "Slide part cannot be null.");

            // Find the title placeholder
            PlaceholderShape titlePlaceholder = slidePart.Slide.Descendants<PlaceholderShape>().FirstOrDefault(ph => ph.Type == PlaceholderValues.Title);
            if (titlePlaceholder == null)
                throw new Exception("Couldn't find the title placeholder.");

            // Get the title shape from the placeholder
            DocumentFormat.OpenXml.Presentation.Shape titleShape = titlePlaceholder.Ancestors<DocumentFormat.OpenXml.Presentation.Shape>().FirstOrDefault();
            if (titleShape == null || titleShape.TextBody == null)
                throw new Exception("Couldn't find the title shape or its text body.");

            // Find the paragraph in the title shape
            Paragraph titleParagraph = titleShape.TextBody.Descendants<Paragraph>().FirstOrDefault();
            if (titleParagraph != null)
            {
                var textElement = titleParagraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault();
                if (textElement != null)
                {
                    // Replace the existing title text
                    textElement.Parent.ReplaceChild(new DocumentFormat.OpenXml.Drawing.Text("Output Slide"), textElement);
                }
                else
                {
                    // If there's no existing text element, append a new one
                    titleParagraph.AppendChild(new Run(new DocumentFormat.OpenXml.Drawing.Text("Output Slide")));
                }

                // Ensure alignment and font properties
                titleParagraph.ParagraphProperties = titleParagraph.ParagraphProperties ?? new DocumentFormat.OpenXml.Drawing.ParagraphProperties();
                titleParagraph.ParagraphProperties.Alignment = TextAlignmentTypeValues.Center;

                // Update run properties for the title
                foreach (Run run in titleParagraph.Descendants<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>();
                    if (runProperties == null)
                    {
                        runProperties = new RunProperties();
                        run.PrependChild(runProperties);
                    }

                    LatinFont latinFont = runProperties.GetFirstChild<LatinFont>();
                    if (latinFont == null)
                    {
                        latinFont = new LatinFont() { Typeface = "Beirut" };
                        runProperties.AppendChild(latinFont);
                    }
                    else
                    {
                        latinFont.Typeface = "Beirut";
                    }
                }
            }
            else
            {
                throw new Exception("Couldn't find a paragraph in the title shape.");
            }
        }

        /// <summary>
        /// Transfers the text from textboxes to corresponding shapes on a slide.
        /// </summary>
        /// <param name="slidePart">The slide part to process.</param>
        /// <exception cref="ArgumentNullException">Thrown when the slidePart is null.</exception>
        public void TransferTextFromTextBoxesToShapes(SlidePart slidePart)
        {
            if (slidePart == null)
            {
                throw new ArgumentNullException(nameof(slidePart), "Slide part cannot be null.");
            }

            // Retrieve text boxes and target shapes from the slide
            var textBoxes = GetTextBoxes(slidePart) ?? throw new Exception("Failed to retrieve text boxes.");
            var textBoxesCopy = textBoxes.ToList();
            var targetShapes = GetTargetShapes(slidePart) ?? throw new Exception("Failed to retrieve target shapes.");

            // Transfer text from each textbox to its corresponding shape
            foreach (var textBox in textBoxesCopy)
            {
                var textBoxBounds = GetTextBoxBounds(textBox);

                if (textBoxBounds == null)
                {
                    throw new Exception("Failed to retrieve bounds for a textbox.");
                }

                // Find a shape that overlaps with the textbox
                var correspondingShape = targetShapes
                    .FirstOrDefault(s =>
                    {
                        var shapeBounds = GetShapeBounds(s);
                        if (shapeBounds == null)
                        {
                            return false;
                        }

                // Check if the shape is below the textbox using their bounds
                        bool isOverlapping =
                        (textBoxBounds.Left <= shapeBounds.Left + shapeBounds.Width) &&
                        (textBoxBounds.Left + textBoxBounds.Width >= shapeBounds.Left) &&
                        (textBoxBounds.Top <= shapeBounds.Top + shapeBounds.Height) &&
                        (textBoxBounds.Top + textBoxBounds.Height >= shapeBounds.Top);

                        return isOverlapping;
                    });

                // Transfer text from the textbox to the shape
                if (correspondingShape != null)
                {
                    TransferText(textBox, correspondingShape);
                    textBox.Remove();
                }
            }

            slidePart.Slide.Save();
        }

        /// <summary>
        /// Retrieves text boxes from a given slide part.
        /// </summary>
        /// <param name="slidePart">The slide part to process.</param>
        /// <returns>An IEnumerable of Shape objects representing text boxes.</returns>
        /// <exception cref="ArgumentNullException">Thrown when the slidePart or its slide is null.</exception>
        private IEnumerable<DocumentFormat.OpenXml.Presentation.Shape> GetTextBoxes(SlidePart slidePart)
        {
            if (slidePart == null || slidePart.Slide == null)
            {
                throw new ArgumentNullException(nameof(slidePart), "Slide part or the slide within the slide part cannot be null.");
            }

            // Assuming text boxes are identified by some criteria, adjust as necessary
            IEnumerable<DocumentFormat.OpenXml.Presentation.Shape> textBoxes = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().Where(IsTextbox);
            return textBoxes ?? Enumerable.Empty<DocumentFormat.OpenXml.Presentation.Shape>();
        }

        /// <summary>
        /// Retrieves target shapes from a given slide part.
        /// </summary>
        /// <param name="slidePart">The slide part to process.</param>
        /// <returns>An IEnumerable of Shape objects representing target shapes.</returns>
        /// <exception cref="ArgumentNullException">Thrown when the slidePart or its slide is null.</exception>
        private IEnumerable<DocumentFormat.OpenXml.Presentation.Shape> GetTargetShapes(SlidePart slidePart)
        {
            if (slidePart == null || slidePart.Slide == null)
            {
                throw new ArgumentNullException(nameof(slidePart), "Slide part or the slide within the slide part cannot be null.");
            }

            // This function fetches the target shapes (which can be Chevrons or Pentagons in this context).
            IEnumerable<DocumentFormat.OpenXml.Presentation.Shape> shapes = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().Where(IsShape);
            return shapes ?? Enumerable.Empty<DocumentFormat.OpenXml.Presentation.Shape>();
        }

        /// <summary>
        /// Transfers text from a source shape to a target shape.
        /// </summary>
        /// <param name="source">The source shape containing the text to transfer.</param>
        /// <param name="target">The target shape to receive the text.</param>
        /// <exception cref="ArgumentNullException">Thrown when the source or target shape is null.</exception>
        /// <exception cref="ArgumentException">Thrown when the source shape does not contain a text body.</exception>
        private void TransferText(DocumentFormat.OpenXml.Presentation.Shape source, DocumentFormat.OpenXml.Presentation.Shape target)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source), "Source shape cannot be null.");
            }

            if (source.TextBody == null)
            {
                throw new ArgumentException("Source shape does not contain a text body.", nameof(source));
            }

            if (target == null)
            {
                throw new ArgumentNullException(nameof(target), "Target shape cannot be null.");
            }

            // Replace or update the target shape's text body with the source's text
            if (target.TextBody == null)
            {
                target.TextBody = new DocumentFormat.OpenXml.Presentation.TextBody();
            }

            target.TextBody.RemoveAllChildren<Paragraph>();
            foreach (var paragraph in source.TextBody.Descendants<Paragraph>().ToList())
            {
                target.TextBody.Append(paragraph.CloneNode(true));
            }
        }

        private dynamic GetShapeBounds(DocumentFormat.OpenXml.Presentation.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape), "Shape cannot be null.");
            }

            if (shape.ShapeProperties?.Transform2D != null)
            {
                return new
                {
                    Top = shape.ShapeProperties.Transform2D.Offset.Y.Value,
                    Left = shape.ShapeProperties.Transform2D.Offset.X.Value,
                    Width = shape.ShapeProperties.Transform2D.Extents.Cx.Value,
                    Height = shape.ShapeProperties.Transform2D.Extents.Cy.Value
                };
            }

            return null;
        }


        /// <summary>
        /// Retrieves the bounds of the provided TextBox shape.
        /// </summary>
        /// <param name="shape">The TextBox shape whose bounds are to be retrieved.</param>
        /// <returns>A dynamic object representing the bounds (Top, Left, Width, Height) of the TextBox.</returns>
        /// <exception cref="ArgumentNullException">Thrown when the provided shape is null.</exception>
        /// <exception cref="ArgumentException">Thrown when the provided shape is not a TextBox.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the shape does not have transformation properties.</exception>
        private dynamic GetTextBoxBounds(DocumentFormat.OpenXml.Presentation.Shape shape)
        {
            // Ensure shape is not null
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape), "Shape cannot be null.");
            }

            // Ensure shape is a TextBox
            if (!IsTextbox(shape))
            {
                throw new ArgumentException("Provided shape is not a TextBox.", nameof(shape));
            }

            // Ensure shape has transformation properties
            if (shape.ShapeProperties?.Transform2D == null)
            {
                throw new InvalidOperationException("Shape does not have transformation properties.");
            }

            // Return the bounds of the shape
            return new
            {
                Top = shape.ShapeProperties.Transform2D.Offset.Y.Value,
                Left = shape.ShapeProperties.Transform2D.Offset.X.Value,
                Width = shape.ShapeProperties.Transform2D.Extents.Cx.Value,
                Height = shape.ShapeProperties.Transform2D.Extents.Cy.Value
            };
        }

        /// <summary>
        /// Determines if the provided shape is a TextBox.
        /// </summary>
        /// <param name="shape">The shape to be checked.</param>
        /// <returns>True if the shape is a TextBox, otherwise false.</returns>
        /// <exception cref="ArgumentNullException">Thrown when the provided shape is null.</exception>
        private static bool IsTextbox(DocumentFormat.OpenXml.Presentation.Shape shape)
        {
            // Ensure shape is not null
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape), "Shape cannot be null.");
            }

            // Check if the name of the shape contains the word "TextBox"
            var shapeName = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.ToString();
            return shapeName?.Contains("TextBox") == true;
        }

        /// <summary>
        /// Aligns and resizes shapes (such as Chevrons) on the given slide part.
        /// This method assumes that the desired positioning and sizing are based on the first chevron shape found.
        /// </summary>
        /// <param name="slidePart">The slide part containing shapes to be aligned and resized.</param>
        /// <exception cref="ArgumentNullException">Thrown when the slidePart or its content is null.</exception>
        /// <exception cref="InvalidOperationException">Thrown when shapes do not have transformation or text properties.</exception>
        public void AlignAndResizeShapes(SlidePart slidePart)
        {
            // Ensure slidePart and its content are not null
            if (slidePart == null || slidePart.Slide == null)
            {
                throw new ArgumentNullException(nameof(slidePart), "Slide part or its content is null.");
            }

            // Retrieve all chevron shapes on the slide, sorted by their horizontal position
            var chevronShapes = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>()
                                   .Where(IsShape)
                                   .OrderBy(shape => shape.ShapeProperties.Transform2D?.Offset.X.Value ?? double.MaxValue)
                                   .ToList();

            // Exit if no chevron shapes found
            if (!chevronShapes.Any())
            {
                Console.WriteLine("No matching shapes found.");
                return;
            }

            // Ensure the first shape has transformation properties
            var firstShape = chevronShapes.First();
            if (firstShape.ShapeProperties.Transform2D == null)
            {
                throw new InvalidOperationException("First shape does not have transformation properties.");
            }

            // Define the desired size and position based on the first shape
            var desiredWidth = firstShape.ShapeProperties.Transform2D.Extents.Cx.Value;
            var desiredHeight = firstShape.ShapeProperties.Transform2D.Extents.Cy.Value;
            var desiredTop = firstShape.ShapeProperties.Transform2D.Offset.Y.Value;
            long currentXOffset = firstShape.ShapeProperties.Transform2D.Offset.X.Value;

            foreach (var shape in chevronShapes)
            {
                // Skip shapes without transformation properties
                if (shape.ShapeProperties.Transform2D == null)
                    continue;

                // Set the desired size
                shape.ShapeProperties.Transform2D.Extents.Cx.Value = (long)(3.04 * 914400);
                shape.ShapeProperties.Transform2D.Extents.Cy.Value = (long)(1.58 * 914400);

                // Set the desired position
                shape.ShapeProperties.Transform2D.Offset.Y.Value = desiredTop;
                shape.ShapeProperties.Transform2D.Offset.X.Value = currentXOffset;

                // Adjust the horizontal offset for the next shape
                currentXOffset += desiredWidth + 150000;

                // Ensure shape has text properties
                if (shape.TextBody == null)
                {
                    throw new InvalidOperationException("Shape does not contain a TextBody.");
                }

                // Center-align text horizontally
                foreach (var paragraph in shape.TextBody.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                {
                    if (paragraph.ParagraphProperties == null)
                    {
                        paragraph.PrependChild(new DocumentFormat.OpenXml.Drawing.ParagraphProperties());
                    }
                    paragraph.ParagraphProperties.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center;
                }

                // Center-align text vertically
                if (shape.TextBody.BodyProperties == null)
                {
                    shape.TextBody.PrependChild(new DocumentFormat.OpenXml.Drawing.BodyProperties());
                }
                shape.TextBody.BodyProperties.Anchor = DocumentFormat.OpenXml.Drawing.TextAnchoringTypeValues.Center;
            }
        }

        /// <summary>
        /// Determines if the provided shape is either a Chevron or Pentagon.
        /// </summary>
        /// <param name="shape">The shape to be checked.</param>
        /// <returns>True if the shape is a Chevron or Pentagon; otherwise false.</returns>
        private bool IsShape(DocumentFormat.OpenXml.Presentation.Shape shape)
        {
            // Return false if shape is null
            if (shape == null)
            {
                return false;
            }

            // Extract the shape name
            var shapeName = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.ToString();

            // Check if the shape name contains "Chevron" or "Pentagon"
            return shapeName?.Contains("Chevron") == true || shapeName?.Contains("Pentagon") == true;
        }

        /// <summary>
        /// Changes the font of all text boxes on a slide to "Beirut".
        /// </summary>
        /// <param name="slidePart">The slide part to be processed.</param>
        public void ChangeFontToBeirut(SlidePart slidePart)
        {
            // Check for null or missing slide content
            if (slidePart == null || slidePart.Slide == null)
            {
                throw new ArgumentNullException(nameof(slidePart), "Slide part or its content is null.");
            }

            // Iterate through all text boxes in the slide
            foreach (var textBox in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>())
            {
                // Skip if there's no TextBody
                if (textBox.TextBody == null)
                {
                    continue;
                }

                // Iterate through all runs in the TextBody
                foreach (var run in textBox.TextBody.Descendants<Run>())
                {
                    RunProperties runProperties = run.GetFirstChild<RunProperties>();
                    if (runProperties == null)
                    {
                        runProperties = new RunProperties();
                        run.PrependChild(runProperties);
                    }

                    // Set the font to "Beirut"
                    LatinFont latinFont = runProperties.GetFirstChild<LatinFont>();
                    if (latinFont == null)
                    {
                        latinFont = new LatinFont() { Typeface = "Beirut" };
                        runProperties.AppendChild(latinFont);
                    }
                    else
                    {
                        latinFont.Typeface = "Beirut";
                    }
                }
            }
        }

        /// <summary>
        /// Aligns and resizes all text boxes on a slide based on the properties of the first text box.
        /// </summary>
        /// <param name="slidePart">The slide part to be processed.</param>
        public void AlignAndResizeTextboxes(SlidePart slidePart)
        {
            // Check for null or missing slide content
            if (slidePart == null || slidePart.Slide == null)
            {
                throw new ArgumentNullException(nameof(slidePart), "Slide part or its content is null.");
            }

            // Retrieve all text boxes on the slide, sorted by their X offset
            var textboxShapes = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>()
                                .Where(IsTextbox)
                                .OrderBy(shape => shape.ShapeProperties?.Transform2D?.Offset?.X?.Value ?? 0)
                                .ToList();

            if (!textboxShapes.Any())
                return;

            // Extract properties from the first text box
            var desiredWidth = textboxShapes.First().ShapeProperties.Transform2D.Extents.Cx.Value;
            var desiredHeight = textboxShapes.First().ShapeProperties.Transform2D.Extents.Cy.Value;
            var desiredTop = textboxShapes.First().ShapeProperties.Transform2D.Offset.Y.Value;
            long currentXOffset = textboxShapes.First().ShapeProperties.Transform2D.Offset.X.Value;

            // Align and resize each text box
            foreach (var shape in textboxShapes)
            {
                if (shape.ShapeProperties?.Transform2D == null)
                    continue;

                // Set size and position
                shape.ShapeProperties.Transform2D.Extents.Cx.Value = desiredWidth;
                shape.ShapeProperties.Transform2D.Extents.Cy.Value = desiredHeight;
                shape.ShapeProperties.Transform2D.Offset.Y.Value = desiredTop;
                shape.ShapeProperties.Transform2D.Offset.X.Value = currentXOffset;

                // Move to the next position
                currentXOffset += desiredWidth + 800000;  // Add margin between shapes
            }
        }

        /// <summary>
        /// Changes the bullet style of paragraphs on a slide to dots.
        /// </summary>
        /// <param name="slidePart">The slide part to be processed.</param>
        public void ChangeBulletToDot(SlidePart slidePart)
        {
            // Check for null or missing slide content
            if (slidePart == null || slidePart.Slide == null)
            {
                throw new ArgumentNullException(nameof(slidePart), "Slide part or its content is null.");
            }

            // Process each paragraph in the slide
            foreach (var paragraph in slidePart.Slide.Descendants<Paragraph>())
            {
                if (paragraph.ParagraphProperties == null)
                {
                    paragraph.ParagraphProperties = new ParagraphProperties();
                }

                // Check for bullet points
                if (HasBulletPoints(paragraph) || HasNumberedBulletPoints(paragraph))
                {
                    // Define the bullet style
                    var bulletFont = new BulletFont { Typeface = "Arial" };
                    var bulletChar = new CharacterBullet { Char = "•" };

                    // Remove numbered bullet if it exists
                    var autoNumberedBullet = paragraph.ParagraphProperties.GetFirstChild<AutoNumberedBullet>();
                    if (autoNumberedBullet != null)
                    {
                        paragraph.ParagraphProperties.RemoveChild(autoNumberedBullet);
                    }

                    // Update or add the bullet style
                    var existingBulletFont = paragraph.ParagraphProperties.GetFirstChild<BulletFont>();
                    if (existingBulletFont == null)
                    {
                        paragraph.ParagraphProperties.Append(bulletFont);
                    }
                    else
                    {
                        existingBulletFont.Typeface = "Arial";
                    }

                    var existingBulletChar = paragraph.ParagraphProperties.GetFirstChild<CharacterBullet>();
                    if (existingBulletChar == null)
                    {
                        paragraph.ParagraphProperties.Append(bulletChar);
                    }
                    else
                    {
                        existingBulletChar.Char = "•";
                    }
                }
            }
        }
        /// <summary>
        /// Determines if the given paragraph has bullet points.
        /// </summary>
        /// <param name="paragraph">The paragraph to check.</param>
        /// <returns><c>true</c> if the paragraph has bullet points; otherwise, <c>false</c>.</returns>
        /// <exception cref="ArgumentNullException">Thrown when the given paragraph is null.</exception>
        private bool HasBulletPoints(Paragraph paragraph)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph), "Paragraph is null.");
            }

            return paragraph.ParagraphProperties?.GetFirstChild<CharacterBullet>() != null;
        }

        /// <summary>
        /// Determines if the given paragraph has numbered bullet points.
        /// </summary>
        /// <param name="paragraph">The paragraph to check.</param>
        /// <returns><c>true</c> if the paragraph has numbered bullet points; otherwise, <c>false</c>.</returns>
        /// <exception cref="ArgumentNullException">Thrown when the given paragraph is null.</exception>
        private bool HasNumberedBulletPoints(Paragraph paragraph)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph), "Paragraph is null.");
            }

            return paragraph.ParagraphProperties?.GetFirstChild<AutoNumberedBullet>() != null;
        }

        /// <summary>
        /// Removes bold and underline styling from text in the given slide part.
        /// </summary>
        /// <param name="slidePart">The slide part to modify.</param>
        /// <exception cref="ArgumentNullException">Thrown when the given slide part or its content is null.</exception>
        public void RemoveBoldAndUnderline(SlidePart slidePart)
        {
            if (slidePart == null || slidePart.Slide == null)
            {
                throw new ArgumentNullException(nameof(slidePart), "Slide part or its content is null.");
            }

            // Iterate through text boxes and remove bold and underline properties from each run of text.
            foreach (var textBox in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().Where(IsTextbox))
            {
                foreach (var run in textBox.TextBody.Descendants<Run>())
                {
                    if (run.RunProperties != null)
                    {
                        run.RunProperties.Bold = null;
                        run.RunProperties.Underline = null;
                    }
                }
            }
        }

        /// <summary>
        /// Logs the provided error message to a specified log file.
        /// </summary>
        /// <param name="errorMessage">The error message to log.</param>
        private void LogError(string errorMessage)
        {
            // Define the path to the error log file.
            string logFilePath = "AppErrorLog.txt";

            // Append the error message with a timestamp to the log file.
            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                writer.WriteLine(DateTime.Now + ": " + errorMessage);
            }
        }

        /// <summary>
        /// Event handler for a button click event. Processes the chosen PowerPoint slide.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The event arguments.</param>
        private void btnProcess_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PowerPoint Files|*.pptx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(openFileDialog.FileName) || !File.Exists(openFileDialog.FileName))
                        {
                            MessageBox.Show("Invalid file path or file does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Process the chosen PowerPoint slide.
                        using (PresentationDocument presentationDocument = PresentationDocument.Open(openFileDialog.FileName, true))
                        {
                            SlideId slideId = presentationDocument.PresentationPart.Presentation.SlideIdList.GetFirstChild<SlideId>();
                            if (slideId == null)
                            {
                                MessageBox.Show("No slides found in the presentation.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }

                            SlidePart slidePart = (SlidePart)presentationDocument.PresentationPart.GetPartById(slideId.RelationshipId);
                            if (slidePart != null)
                            {
                                UpdateTitle(slidePart);
                                TransferTextFromTextBoxesToShapes(slidePart);
                                AlignAndResizeShapes(slidePart);
                                ChangeFontToBeirut(slidePart);
                                AlignAndResizeTextboxes(slidePart);
                                ChangeBulletToDot(slidePart);
                                RemoveBoldAndUnderline(slidePart);
                            }
                        }

                        label.Text = "          Slide Updated Successfully!";
                        label.TextAlign = ContentAlignment.MiddleCenter;
                    }
                    catch (Exception ex)
                    {
                        LogError(ex.Message);
                        label.Text = $"An error occurred: {ex.Message}";
                        MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    label.Text = "Operation cancelled by the user.";
                    label.TextAlign = ContentAlignment.MiddleCenter;
                }
            }
        }
    }
}
