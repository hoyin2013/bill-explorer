interface Props {
  message: string
  onClear: () => void
}

export function ErrorMessage({ message, onClear }: Props) {
  if (!message) return null
  return (
    <div className="error-msg" onClick={onClear}>
      <span className="error-icon">!</span>
      <span>{message}</span>
      <span className="error-close">点击关闭</span>
    </div>
  )
}
