export default function Modal({ title, onClose, children, maxWidth = 'max-w-2xl' }) {
  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
      <div className={`bg-white shadow-xl w-full ${maxWidth} mx-4`}>
        <div className="flex items-center justify-between px-6 py-3 border-b border-[#DEE2E6] bg-[#F8F8F8]">
          <h2 className="text-sm font-semibold text-[#333]">{title}</h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-700 text-xl font-bold leading-none">×</button>
        </div>
        {children}
      </div>
    </div>
  );
}
