interface EmptyStateProps {
  title: string;
  description?: string;
}

export default function EmptyState({ title, description }: EmptyStateProps) {
  return (
    <div className="bg-white border border-border-light p-8 text-center">
      <h3 className="text-sm font-semibold text-[#333] mb-1">{title}</h3>
      {description && <p className="text-xs text-text-secondary">{description}</p>}
    </div>
  );
}
