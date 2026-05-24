"use client"

import * as React from "react"
import { CheckIcon, ChevronDownIcon } from "lucide-react"

import { cn } from "@/lib/utils"

type SelectContextValue = {
  readonly open: boolean
  readonly selectedLabel: React.ReactNode
  readonly value: string
  readonly setOpen: (open: boolean) => void
  readonly onValueChange: (value: string) => void
  readonly registerItem: (value: string, label: React.ReactNode) => void
}

const SelectContext = React.createContext<SelectContextValue | null>(null)

const useSelectContext = (): SelectContextValue => {
  const context = React.useContext(SelectContext)
  if (context === null) {
    throw new Error("Select components must be used inside <Select>.")
  }

  return context
}

type SelectProps = {
  readonly children: React.ReactNode
  readonly value: string
  readonly onValueChange: (value: string) => void
}

function Select({ children, value, onValueChange }: SelectProps) {
  const [open, setOpen] = React.useState(false)
  const [items, setItems] = React.useState<ReadonlyMap<string, React.ReactNode>>(
    () => new Map()
  )

  const registerItem = React.useCallback(
    (itemValue: string, label: React.ReactNode): void => {
      setItems((currentItems) => {
        if (currentItems.get(itemValue) === label) {
          return currentItems
        }

        const nextItems = new Map(currentItems)
        nextItems.set(itemValue, label)
        return nextItems
      })
    },
    []
  )

  const contextValue = React.useMemo<SelectContextValue>(
    () => ({
      open,
      selectedLabel: items.get(value) ?? value,
      value,
      setOpen,
      onValueChange,
      registerItem,
    }),
    [items, onValueChange, open, registerItem, value]
  )

  return (
    <SelectContext.Provider value={contextValue}>
      <div className="relative" data-slot="select">
        {children}
      </div>
    </SelectContext.Provider>
  )
}

function SelectValue({ className }: { readonly className?: string }) {
  const { selectedLabel } = useSelectContext()

  return (
    <span className={cn("flex flex-1 text-left", className)} data-slot="select-value">
      {selectedLabel}
    </span>
  )
}

type SelectTriggerProps = React.ButtonHTMLAttributes<HTMLButtonElement> & {
  readonly size?: "sm" | "default"
}

function SelectTrigger({
  className,
  size = "default",
  children,
  type = "button",
  ...props
}: SelectTriggerProps) {
  const { open, setOpen } = useSelectContext()

  return (
    <button
      aria-expanded={open}
      className={cn(
        "flex w-fit items-center justify-between gap-1.5 rounded-lg border border-input bg-transparent py-2 pr-2 pl-2.5 text-sm whitespace-nowrap transition-colors outline-none select-none focus-visible:border-ring focus-visible:ring-3 focus-visible:ring-ring/50 disabled:cursor-not-allowed disabled:opacity-50 data-[size=default]:h-8 data-[size=sm]:h-7 data-[size=sm]:rounded-[min(var(--radius-md),10px)] dark:bg-input/30",
        className
      )}
      data-size={size}
      data-slot="select-trigger"
      onClick={(event) => {
        props.onClick?.(event)
        if (!event.defaultPrevented) {
          setOpen(!open)
        }
      }}
      type={type}
      {...props}
    >
      {children}
      <ChevronDownIcon className="pointer-events-none size-4 text-muted-foreground" />
    </button>
  )
}

function SelectContent({
  className,
  children,
}: {
  readonly className?: string
  readonly children: React.ReactNode
}) {
  const { open } = useSelectContext()

  if (!open) {
    return null
  }

  return (
    <div
      className={cn(
        "absolute z-50 mt-2 max-h-64 w-full min-w-36 overflow-y-auto rounded-lg bg-popover p-1 text-popover-foreground shadow-md ring-1 ring-foreground/10",
        className
      )}
      data-slot="select-content"
    >
      {children}
    </div>
  )
}

function SelectItem({
  className,
  children,
  value,
}: {
  readonly className?: string
  readonly children: React.ReactNode
  readonly value: string
}) {
  const { value: selectedValue, onValueChange, registerItem, setOpen } = useSelectContext()

  React.useEffect(() => {
    registerItem(value, children)
  }, [children, registerItem, value])

  const isSelected = selectedValue === value

  return (
    <button
      className={cn(
        "relative flex w-full cursor-default items-center gap-1.5 rounded-md py-1.5 pr-8 pl-2 text-left text-sm outline-hidden select-none hover:bg-accent hover:text-accent-foreground focus:bg-accent focus:text-accent-foreground",
        className
      )}
      data-slot="select-item"
      onClick={() => {
        onValueChange(value)
        setOpen(false)
      }}
      type="button"
    >
      <span className="flex flex-1 shrink-0 gap-2 whitespace-nowrap">{children}</span>
      {isSelected ? (
        <span className="pointer-events-none absolute right-2 flex size-4 items-center justify-center">
          <CheckIcon className="size-4" />
        </span>
      ) : null}
    </button>
  )
}

function SelectGroup({
  className,
  ...props
}: React.HTMLAttributes<HTMLDivElement>) {
  return <div className={cn("scroll-my-1 p-1", className)} data-slot="select-group" {...props} />
}

function SelectLabel({
  className,
  ...props
}: React.HTMLAttributes<HTMLDivElement>) {
  return (
    <div
      className={cn("px-1.5 py-1 text-xs text-muted-foreground", className)}
      data-slot="select-label"
      {...props}
    />
  )
}

function SelectSeparator({
  className,
  ...props
}: React.HTMLAttributes<HTMLDivElement>) {
  return (
    <div
      className={cn("pointer-events-none -mx-1 my-1 h-px bg-border", className)}
      data-slot="select-separator"
      {...props}
    />
  )
}

function SelectScrollUpButton() {
  return null
}

function SelectScrollDownButton() {
  return null
}

export {
  Select,
  SelectContent,
  SelectGroup,
  SelectItem,
  SelectLabel,
  SelectScrollDownButton,
  SelectScrollUpButton,
  SelectSeparator,
  SelectTrigger,
  SelectValue,
}
