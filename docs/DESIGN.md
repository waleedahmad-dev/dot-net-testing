# Design System - Word Highlighter

## Color Palette

### Primary Colors

- **Brand Green**: `#15C39A` - Main actions, highlights, success states
- **Brand Green Hover**: `#12B38A` - Hover state for green buttons
- **Brand Green Light**: `#F0FDF9` - Light backgrounds, subtle highlights

### Neutral Colors

- **Gray 50**: `#F9FAFB` - Light backgrounds, input backgrounds
- **Gray 100**: `#F3F4F6` - Secondary button backgrounds
- **Gray 200**: `#E5E7EB` - Borders, dividers
- **Gray 300**: `#D1D5DB` - Input borders
- **Gray 400**: `#9CA3AF` - Disabled text
- **Gray 500**: `#6B7280` - Secondary text
- **Gray 600**: `#4B5563` - Body text
- **Gray 900**: `#111827` - Headings, primary text

### Background Colors

- **App Background**: `#F5F7FA` - Main application background
- **Card Background**: `#FFFFFF` - Card and widget backgrounds

### Status Colors

- **Success Green**: `#10B981` - Success messages
- **Warning Yellow**: `#F59E0B` - Warning states, disconnected status
- **Warning Light**: `#FEF3C7` - Warning backgrounds
- **Error Red**: `#DC2626` - Error messages
- **Error Light**: `#FEE2E2` - Error backgrounds
- **Danger**: `#FF6B6B` - Destructive actions

### Accent Colors

- **Blue Light**: `#DBEAFE` - Informational backgrounds
- **Green Light**: `#DCFCE7` - Success backgrounds

## Typography

### Font Families

- **Primary**: Segoe UI (system default)
- **Monospace**: Consolas (for code/logs)

### Font Sizes

- **Heading 1**: 24px - Main page titles
- **Heading 2**: 22px - App title
- **Heading 3**: 20px - Section titles
- **Heading 4**: 18px - Card titles
- **Body Large**: 15px - Input text, important body
- **Body**: 14px - Regular body text, buttons
- **Body Small**: 13px - Secondary text, status
- **Caption**: 12px - Labels, tags

### Font Weights

- **Bold**: 700 - Main headings
- **SemiBold**: 600 - Subheadings, buttons
- **Medium**: 500 - Important body text
- **Regular**: 400 - Body text

## Spacing

### Padding Scale

- **XS**: 8px
- **S**: 12px
- **M**: 16px
- **L**: 20px
- **XL**: 24px
- **2XL**: 28px
- **3XL**: 30px
- **4XL**: 40px

### Margin Scale

Same as padding scale

### Common Spacings

- Card Padding: 30px
- Card Margin: 24px
- Button Padding: 20px 12px
- Input Padding: 14px
- Section Spacing: 20px

## Border Radius

- **Small**: 6px - Quick action buttons
- **Medium**: 8px - Buttons, inputs, small cards
- **Large**: 12px - Icons, badges
- **XLarge**: 16px - Main cards
- **Round**: 20px - Status badges
- **Circle**: 50% - Circular elements

## Shadows

### Card Shadow

```
Color: #000000
Opacity: 0.08
Blur Radius: 20px
Shadow Depth: 0
```

### Widget Shadow

```
Color: #000000
Opacity: 0.15
Blur Radius: 30px
Shadow Depth: 0
```

## Button Styles

### Primary Button (ModernButton)

- Background: #15C39A
- Text: White
- Padding: 20px 12px
- Border Radius: 8px
- Hover: #12B38A

### Secondary Button

- Background: #F3F4F6
- Text: #374151
- Padding: 16px 10px
- Border Radius: 8px
- Hover: #E5E7EB

### Danger Button

- Background: #FF6B6B
- Text: White
- Padding: 20px 12px
- Border Radius: 8px
- Hover: #FF5252

### Quick Action Button

- Background: White
- Border: 1px #E5E7EB
- Text: #6B7280
- Padding: 12px 6px
- Border Radius: 6px
- Hover Border: #15C39A

## Input Styles

### Text Input

- Height: 44px
- Border: 2px solid #D1D5DB
- Border Radius: 8px
- Padding: 14px
- Font Size: 15px
- Focus: Border color changes to primary

## Card Styles

### Standard Card

- Background: White
- Border Radius: 16px
- Padding: 30px
- Shadow: Drop shadow (0.08 opacity, 20px blur)

### Action Card

- Background: White
- Border Radius: 16px
- Padding: 24px
- Shadow: Drop shadow (0.08 opacity, 20px blur)
- Icon Background: Colored circle (56x56px, radius 12px)

## Icon Guidelines

### Icon Sizes

- Small: 24x24px
- Medium: 48x48px
- Large: 56x56px
- XLarge: 80x80px

### Icon Backgrounds

- Highlight Card: #DBEAFE (Blue Light)
- Clear Card: #FEE2E2 (Red Light)
- Success: #F0FDF9 (Green Light)

## Layout

### Window Sizes

- Main Window: 1000x650px
- Replacement Widget: 480x380px

### Grid Spacing

- Column Gap: 24px
- Row Gap: 24px

### Responsive Behavior

- Cards stack vertically on narrow windows
- Minimum window width: 800px

## Interactive States

### Hover States

- Buttons: Slightly darker background
- Cards: No change (already has shadow)
- Links: Underline appears

### Focus States

- Inputs: Border color changes to primary
- Buttons: Subtle outline

### Disabled States

- Background: #E0E0E0
- Text: #9E9E9E
- No hover effect

## Animations

### Transition Duration

- Fast: 150ms
- Normal: 200ms
- Slow: 300ms

### Easing

- Standard: ease-in-out
- Enter: ease-out
- Exit: ease-in

## Accessibility

### Contrast Ratios

- Normal Text: Minimum 4.5:1
- Large Text: Minimum 3:1
- Interactive Elements: Minimum 3:1

### Focus Indicators

- Visible focus outline on all interactive elements
- Keyboard navigation support

## Emojis Used

- 🚀 Application Started
- ✅ Success/Connected
- ❌ Failed/Error
- ⚠️ Warning
- 🎨 Highlighting
- ✨ Highlighted Successfully
- 🧹 Clearing
- ✓ Cleared/Done
- 📝 Replaced
- 👆 Click Detection
- 🔤 Uppercase
- 🔡 Lowercase
- 🔠 Capitalize

## Best Practices

1. **Consistency**: Use defined colors, spacing, and typography consistently
2. **White Space**: Don't crowd elements, use generous padding
3. **Hierarchy**: Clear visual hierarchy with font sizes and weights
4. **Feedback**: Provide immediate visual feedback for all interactions
5. **Simplicity**: Keep UI clean and uncluttered
6. **Professional**: Maintain Grammarly-inspired professional aesthetic
