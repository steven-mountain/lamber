import type { ComponentPropsWithoutRef } from "react"

import { cn } from "@/lib/utils"
import { iconMap, type AppIconName } from "./iconMap"
export type { AppIconName } from "./iconMap"

export interface AppIconProps extends Omit<ComponentPropsWithoutRef<"svg">, "name"> {
  name: AppIconName
  size?: number
  strokeWidth?: number
  title?: string
}

export default function AppIcon({
  name,
  size = 18,
  strokeWidth = 1.75,
  className,
  title,
  "aria-label": ariaLabel,
  "aria-hidden": ariaHidden,
  ...props
}: AppIconProps) {
  const Icon = iconMap[name]
  const isDecorative = ariaHidden ?? (title || ariaLabel ? undefined : true)

  return (
    <Icon
      aria-hidden={isDecorative}
      aria-label={ariaLabel}
      className={cn("shrink-0", className)}
      height={size}
      role={title || ariaLabel ? "img" : undefined}
      strokeWidth={strokeWidth}
      width={size}
      {...props}
    >
      {title ? <title>{title}</title> : null}
    </Icon>
  )
}
