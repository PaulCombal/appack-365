## Installation steps

The MS365 AppPack has been created using the steps below.
This folder was initially created with `appack creator new`, which created an image file `image.qcow2`.
You can recreate the image file manually with: `qemu-img create -f qcow2 image.qcow2 32G`.
You will need to temporarily add `/drive:home,$HOME` to your RDP options to copy files over.

1. Copy your Windows installation medium in AppPack/installer.iso
2. Copy your guest addons ISO in AppPack/guest-addons.iso
3. Run `appack creator boot-install`
4. Install Windows (turn off all settings during installation)
5. Install guest addons from the ISO
6. Press Super+I to open settings
   * Search for "Remote Desktop" and enable it
   * Go to Windows Update and update Windows
   * Disable Shared experiences
   * Disable Notifications
   * Uninstall Windows Hello Face
7. Shut down Windows
8. Run `appack creator boot`
9. Copy all vbs files in guest/C to C:\
10. Load the registry file in guest/registry.reg
11. Install MS 365
12. Uninstall OneDrive and other unnecessary apps if any
13. Restart (disconnect drive mounts if any)
14. Check if there are system restore points?
15. Run `sdelete -z C:` inside the VM
16. Shut down
17. Run `qemu-img convert -O qcow2 image.qcow2 smaller_image.qcow2`
18. Run `mv image.qcow2 image_backup.qcow2`
19. Run `mv smaller_image.qcow2 image.qcow2`
20. Run `appack creator boot`
21. Once Windows is booted and idle, run `appack creator snapshot`

## Areas of improvement

- The animations seem to be disabled when using RDP