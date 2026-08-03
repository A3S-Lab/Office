import { useCallback, useEffect, useState } from 'react';

export function usePlaygroundSidebarState(modal: boolean): {
  closeAll: () => void;
  open: boolean;
  setOpen: (open: boolean) => void;
} {
  const [persistentOpen, setPersistentOpen] = useState(true);
  const [modalOpen, setModalOpen] = useState(false);

  useEffect(() => {
    if (!modal) setModalOpen(false);
  }, [modal]);

  const setOpen = useCallback(
    (open: boolean) => {
      if (modal) setModalOpen(open);
      else setPersistentOpen(open);
    },
    [modal],
  );
  const closeAll = useCallback(() => {
    setPersistentOpen(false);
    setModalOpen(false);
  }, []);

  return {
    closeAll,
    open: modal ? modalOpen : persistentOpen,
    setOpen,
  };
}
