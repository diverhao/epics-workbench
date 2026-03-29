package org.epics.workbench.ui

import com.intellij.ui.ColorUtil
import com.intellij.ui.JBColor
import java.awt.Color
import java.awt.Graphics
import java.awt.Graphics2D
import java.awt.RenderingHints
import java.awt.Rectangle
import javax.swing.AbstractButton
import javax.swing.BorderFactory
import javax.swing.ButtonModel
import javax.swing.JComponent
import javax.swing.JTextArea
import javax.swing.JTextField
import javax.swing.border.Border
import javax.swing.plaf.basic.BasicButtonUI
import javax.swing.plaf.basic.BasicGraphicsUtils

internal data class EpicsWidgetPalette(
  val background: Color,
  val foreground: Color,
  val mutedForeground: Color,
  val panelBackground: Color,
  val inputBackground: Color,
  val headerBackground: Color,
  val selectionBackground: Color,
  val selectionForeground: Color,
  val borderColor: Color,
  val separatorColor: Color,
  val linkForeground: Color,
)

internal fun buildEpicsWidgetPalette(background: Color, foreground: Color): EpicsWidgetPalette {
  val accentColor = JBColor(0x2F6FEB, 0x6CA6FF)
  val isDark = ColorUtil.isDark(background)
  return EpicsWidgetPalette(
    background = background,
    foreground = foreground,
    mutedForeground = ColorUtil.mix(foreground, background, 0.35),
    panelBackground = ColorUtil.mix(background, foreground, if (isDark) 0.08 else 0.035),
    inputBackground = ColorUtil.mix(background, foreground, if (isDark) 0.12 else 0.055),
    headerBackground = ColorUtil.mix(background, foreground, if (isDark) 0.16 else 0.10),
    selectionBackground = ColorUtil.mix(background, accentColor, if (isDark) 0.26 else 0.14),
    selectionForeground = foreground,
    borderColor = ColorUtil.mix(background, foreground, if (isDark) 0.26 else 0.18),
    separatorColor = ColorUtil.mix(background, foreground, if (isDark) 0.18 else 0.11),
    linkForeground = ColorUtil.mix(foreground, accentColor, if (isDark) 0.62 else 0.68),
  )
}

internal fun applyEpicsWidgetButtonStyle(button: AbstractButton, palette: EpicsWidgetPalette) {
  button.setUI(EpicsWidgetButtonUi(palette))
  button.isOpaque = false
  button.isContentAreaFilled = false
  button.isBorderPainted = false
  button.background = palette.panelBackground
  button.foreground = palette.foreground
  button.isFocusPainted = false
  button.border = createEpicsWidgetBoxBorder(palette)
}

internal fun applyEpicsWidgetTextFieldStyle(field: JTextField, palette: EpicsWidgetPalette) {
  field.isOpaque = true
  field.background = palette.inputBackground
  field.foreground = palette.foreground
  field.caretColor = palette.foreground
  field.selectionColor = palette.selectionBackground
  field.selectedTextColor = palette.selectionForeground
  field.border = createEpicsWidgetBoxBorder(palette)
}

internal fun applyEpicsWidgetTextAreaStyle(area: JTextArea, palette: EpicsWidgetPalette) {
  area.isOpaque = true
  area.background = palette.inputBackground
  area.foreground = palette.foreground
  area.caretColor = palette.foreground
  area.selectionColor = palette.selectionBackground
  area.selectedTextColor = palette.selectionForeground
  area.border = BorderFactory.createCompoundBorder(
    BorderFactory.createLineBorder(palette.borderColor),
    BorderFactory.createEmptyBorder(6, 8, 6, 8),
  )
}

internal fun createEpicsWidgetBoxBorder(palette: EpicsWidgetPalette): Border {
  return BorderFactory.createCompoundBorder(
    BorderFactory.createLineBorder(palette.borderColor),
    BorderFactory.createEmptyBorder(5, 10, 5, 10),
  )
}

private class EpicsWidgetButtonUi(
  private val palette: EpicsWidgetPalette,
) : BasicButtonUI() {
  override fun paint(graphics: Graphics, component: JComponent) {
    val button = component as AbstractButton
    val graphics2d = graphics.create() as Graphics2D
    graphics2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON)
    graphics2d.color = resolveBackground(button.model)
    graphics2d.fillRoundRect(0, 0, component.width, component.height, 10, 10)
    graphics2d.dispose()
    super.paint(graphics, component)
  }

  override fun paintText(
    graphics: Graphics,
    button: AbstractButton,
    textRect: Rectangle,
    text: String,
  ) {
    val fontMetrics = graphics.fontMetrics
    graphics.color = if (button.model.isEnabled) palette.foreground else palette.mutedForeground
    BasicGraphicsUtils.drawStringUnderlineCharAt(
      graphics,
      text,
      button.displayedMnemonicIndex,
      textRect.x + getTextShiftOffset(),
      textRect.y + fontMetrics.ascent + getTextShiftOffset(),
    )
  }

  private fun resolveBackground(model: ButtonModel): Color {
    return when {
      !model.isEnabled -> ColorUtil.mix(palette.panelBackground, palette.background, 0.35)
      model.isPressed || model.isSelected -> palette.selectionBackground
      model.isRollover -> ColorUtil.mix(palette.panelBackground, palette.selectionBackground, 0.22)
      else -> palette.panelBackground
    }
  }
}
