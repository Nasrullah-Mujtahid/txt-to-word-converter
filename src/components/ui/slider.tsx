import * as React from "react"

import { cn } from "@/lib/utils"

type SliderProps = Omit<
  React.InputHTMLAttributes<HTMLInputElement>,
  "defaultValue" | "onChange" | "type" | "value"
> & {
  readonly defaultValue?: number | readonly number[]
  readonly onValueChange?: (value: number | readonly number[]) => void
  readonly value?: number | readonly number[]
}

function Slider({
  className,
  defaultValue,
  value,
  min = 0,
  max = 100,
  onValueChange,
  ...props
}: SliderProps) {
  const currentValue = Array.isArray(value)
    ? value[0]
    : typeof value === "number"
      ? value
      : Array.isArray(defaultValue)
        ? defaultValue[0]
        : typeof defaultValue === "number"
          ? defaultValue
          : min

  return (
    <input
      className={cn(
        "h-2 w-full cursor-pointer appearance-none rounded-full bg-muted accent-primary disabled:cursor-not-allowed disabled:opacity-50",
        className
      )}
      data-slot="slider"
      defaultValue={currentValue}
      max={max}
      min={min}
      onChange={(event) => onValueChange?.(Number(event.target.value))}
      type="range"
      value={currentValue}
      {...props}
    />
  )
}

export { Slider }
