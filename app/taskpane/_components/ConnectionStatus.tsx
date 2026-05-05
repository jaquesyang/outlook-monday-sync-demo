type Variant = 'green' | 'yellow' | 'red';

export function ConnectionStatus(props: {
  label: string;
  variant: Variant;
  detail?: string;
  action?: { label: string; onClick: () => void };
}) {
  const dot =
    props.variant === 'green'
      ? 'bg-emerald-500'
      : props.variant === 'yellow'
        ? 'bg-amber-500'
        : 'bg-rose-500';
  return (
    <div className="flex items-center justify-between rounded border border-zinc-200 px-3 py-2">
      <div className="flex items-center gap-2">
        <span className={`h-2 w-2 rounded-full ${dot}`} />
        <div>
          <div className="font-medium">{props.label}</div>
          {props.detail && <div className="text-xs text-zinc-500">{props.detail}</div>}
        </div>
      </div>
      {props.action && (
        <button
          className="rounded bg-zinc-900 px-2 py-1 text-xs text-white"
          onClick={props.action.onClick}
        >
          {props.action.label}
        </button>
      )}
    </div>
  );
}
