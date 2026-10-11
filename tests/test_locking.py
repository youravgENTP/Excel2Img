import pytest

from excel2img.locking import ActiveRunLock, AlreadyRunningError, active_run_is_locked


def test_active_lock_rejects_overlap_and_is_reusable(tmp_path):
    path = tmp_path / "active.lock"
    assert not active_run_is_locked(path)
    with ActiveRunLock(path):
        assert active_run_is_locked(path)
        with pytest.raises(AlreadyRunningError):
            with ActiveRunLock(path):
                pass
    # The file can remain, but the kernel leaves no stale ownership after exit.
    assert not active_run_is_locked(path)
    with ActiveRunLock(path):
        pass
