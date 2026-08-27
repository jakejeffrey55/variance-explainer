import * as React from "react";
import { cva, type VariantProps } from "class-variance-authority";
import { cn } from "@/lib/utils";

const alertVariants = cva(
  "relative w-full rounded-lg border p-4 text-sm [&>svg]:absolute [&>svg]:left-4 [&>svg]:top-4 [&>svg]:size-4 [&>svg~*]:pl-7",
  {
    variants: {
      variant: {
        default: "bg-card text-card-foreground",
        info: "border-primary/20 bg-primary/5 text-foreground [&>svg]:text-primary",
        warning: "border-warning/30 bg-warning-soft text-foreground [&>svg]:text-warning",
        destructive: "border-destructive/30 bg-destructive/5 text-foreground [&>svg]:text-destructive",
        emergency: "border-emergency-border bg-emergency-soft text-foreground [&>svg]:text-emergency",
        success: "border-success/30 bg-success-soft text-foreground [&>svg]:text-success",
      },
    },
    defaultVariants: { variant: "default" },
  },
);

const Alert = React.forwardRef<
  HTMLDivElement,
  React.HTMLAttributes<HTMLDivElement> & VariantProps<typeof alertVariants>
>(({ className, variant, ...props }, ref) => (
  <div ref={ref} role="alert" className={cn(alertVariants({ variant }), className)} {...props} />
));
Alert.displayName = "Alert";

const AlertTitle = React.forwardRef<HTMLParagraphElement, React.HTMLAttributes<HTMLHeadingElement>>(
  ({ className, ...props }, ref) => (
    <h5 ref={ref} className={cn("mb-1 font-medium leading-none tracking-tight", className)} {...props} />
  ),
);
AlertTitle.displayName = "AlertTitle";

const AlertDescription = React.forwardRef<HTMLParagraphElement, React.HTMLAttributes<HTMLParagraphElement>>(
  ({ className, ...props }, ref) => (
    <div ref={ref} className={cn("text-sm text-muted-foreground [&_p]:leading-relaxed", className)} {...props} />
  ),
);
AlertDescription.displayName = "AlertDescription";

export { Alert, AlertTitle, AlertDescription };
