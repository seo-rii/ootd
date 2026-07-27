const COMPOUND_FILE_SIGNATURE: [u8; 8] = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const HEADER_DIFAT_ENTRIES: usize = 109;
const DIRECTORY_ENTRY_SIZE: usize = 128;
const FREE_SECTOR: u32 = 0xffff_ffff;
const END_OF_CHAIN: u32 = 0xffff_fffe;
const FAT_SECTOR: u32 = 0xffff_fffd;
const DIFAT_SECTOR: u32 = 0xffff_fffc;

pub(crate) fn is_encrypted_ooxml_compound_file(input: &[u8]) -> bool {
    CompoundFileView::parse(input)
        .is_some_and(|compound_file| compound_file.has_encrypted_ooxml_streams())
}

struct CompoundFileView<'a> {
    input: &'a [u8],
    major_version: u16,
    sector_size: usize,
    sector_count: usize,
    first_directory_sector: u32,
    declared_directory_sector_count: Option<usize>,
    fat_sector_ids: Vec<u32>,
}

struct DirectoryEntry {
    object_type: u8,
    name: Option<String>,
    left_sibling: u32,
    right_sibling: u32,
    child: u32,
    start_sector: u32,
    stream_size: u64,
}

impl<'a> CompoundFileView<'a> {
    fn parse(input: &'a [u8]) -> Option<Self> {
        if input.get(..COMPOUND_FILE_SIGNATURE.len())? != COMPOUND_FILE_SIGNATURE {
            return None;
        }
        let major_version = read_u16(input, 26)?;
        let sector_shift = read_u16(input, 30)?;
        if read_u16(input, 28)? != 0xfffe
            || read_u16(input, 32)? != 6
            || input.get(34..40)? != [0; 6]
            || read_u32(input, 56)? != 4096
            || !matches!((major_version, sector_shift), (3, 9) | (4, 12))
        {
            return None;
        }
        let sector_size = 1_usize.checked_shl(u32::from(sector_shift))?;
        if input.len() < sector_size || (input.len() - sector_size) % sector_size != 0 {
            return None;
        }
        let sector_count = (input.len() - sector_size) / sector_size;
        let declared_directory_sector_count = usize::try_from(read_u32(input, 40)?).ok()?;
        let fat_sector_count = usize::try_from(read_u32(input, 44)?).ok()?;
        let first_directory_sector = read_u32(input, 48)?;
        let first_difat_sector = read_u32(input, 68)?;
        let difat_sector_count = usize::try_from(read_u32(input, 72)?).ok()?;
        if sector_count == 0
            || fat_sector_count == 0
            || fat_sector_count > sector_count
            || difat_sector_count > sector_count
            || !is_regular_sector(first_directory_sector, sector_count)
            || (major_version == 3 && declared_directory_sector_count != 0)
            || (major_version == 4
                && !(1..=sector_count).contains(&declared_directory_sector_count))
        {
            return None;
        }

        let mut fat_sector_ids = Vec::with_capacity(fat_sector_count);
        for index in 0..HEADER_DIFAT_ENTRIES {
            let sector = read_u32(input, 76 + index * size_of::<u32>())?;
            if sector != FREE_SECTOR {
                fat_sector_ids.push(sector);
            }
        }

        let entries_per_difat_sector = sector_size.checked_div(size_of::<u32>())?.checked_sub(1)?;
        let mut difat_sector = first_difat_sector;
        let mut visited_difat = vec![false; sector_count];
        for _ in 0..difat_sector_count {
            if !is_regular_sector(difat_sector, sector_count) {
                return None;
            }
            let difat_index = usize::try_from(difat_sector).ok()?;
            if visited_difat[difat_index] {
                return None;
            }
            visited_difat[difat_index] = true;
            let sector = sector_bytes(input, sector_size, difat_sector)?;
            for index in 0..entries_per_difat_sector {
                let fat_sector = read_u32(sector, index * size_of::<u32>())?;
                if fat_sector != FREE_SECTOR {
                    fat_sector_ids.push(fat_sector);
                }
            }
            difat_sector = read_u32(sector, entries_per_difat_sector * size_of::<u32>())?;
        }
        if (difat_sector_count == 0 && first_difat_sector != END_OF_CHAIN)
            || (difat_sector_count > 0 && difat_sector != END_OF_CHAIN)
            || fat_sector_ids.len() != fat_sector_count
        {
            return None;
        }

        let mut seen_fat = vec![false; sector_count];
        for sector in &fat_sector_ids {
            if !is_regular_sector(*sector, sector_count) {
                return None;
            }
            let index = usize::try_from(*sector).ok()?;
            if seen_fat[index] {
                return None;
            }
            seen_fat[index] = true;
        }

        Some(Self {
            input,
            major_version,
            sector_size,
            sector_count,
            first_directory_sector,
            declared_directory_sector_count: (major_version == 4)
                .then_some(declared_directory_sector_count),
            fat_sector_ids,
        })
    }

    fn has_encrypted_ooxml_streams(&self) -> bool {
        let Some(entries) = self.directory_entries() else {
            return false;
        };
        let Some(root) = entries.first() else {
            return false;
        };
        if root.object_type != 5
            || root
                .name
                .as_deref()
                .is_none_or(|name| !name.eq_ignore_ascii_case("Root Entry"))
            || entries.iter().skip(1).any(|entry| entry.object_type == 5)
        {
            return false;
        }

        let mut saw_encryption_info = false;
        let mut saw_encrypted_package = false;
        let mut visited = vec![false; entries.len()];
        let mut pending = Vec::new();
        if root.child != FREE_SECTOR {
            pending.push(root.child);
        }
        while let Some(entry_id) = pending.pop() {
            let Ok(entry_index) = usize::try_from(entry_id) else {
                return false;
            };
            let Some(entry) = entries.get(entry_index) else {
                return false;
            };
            if visited[entry_index] || !matches!(entry.object_type, 1 | 2) {
                return false;
            }
            visited[entry_index] = true;
            if entry.left_sibling != FREE_SECTOR {
                pending.push(entry.left_sibling);
            }
            if entry.right_sibling != FREE_SECTOR {
                pending.push(entry.right_sibling);
            }
            if entry.object_type != 2 {
                continue;
            }
            let Some(name) = entry.name.as_deref() else {
                return false;
            };
            if entry.stream_size == 0
                || matches!(
                    entry.start_sector,
                    FREE_SECTOR | END_OF_CHAIN | FAT_SECTOR | DIFAT_SECTOR
                )
            {
                if name.eq_ignore_ascii_case("EncryptionInfo")
                    || name.eq_ignore_ascii_case("EncryptedPackage")
                {
                    return false;
                }
                continue;
            }
            if name.eq_ignore_ascii_case("EncryptionInfo") {
                if saw_encryption_info {
                    return false;
                }
                saw_encryption_info = true;
            } else if name.eq_ignore_ascii_case("EncryptedPackage") {
                if saw_encrypted_package {
                    return false;
                }
                saw_encrypted_package = true;
            }
        }

        saw_encryption_info && saw_encrypted_package
    }

    fn directory_entries(&self) -> Option<Vec<DirectoryEntry>> {
        let mut entries = Vec::new();
        let mut directory_sector = self.first_directory_sector;
        let mut directory_sector_count = 0;
        let mut visited_directory = vec![false; self.sector_count];

        for _ in 0..self.sector_count {
            if directory_sector == END_OF_CHAIN {
                break;
            }
            if !is_regular_sector(directory_sector, self.sector_count) {
                return None;
            }
            let directory_index = usize::try_from(directory_sector).ok()?;
            if visited_directory[directory_index] {
                return None;
            }
            visited_directory[directory_index] = true;
            directory_sector_count += 1;
            let sector = sector_bytes(self.input, self.sector_size, directory_sector)?;
            for entry in sector.chunks_exact(DIRECTORY_ENTRY_SIZE) {
                let object_type = entry[66];
                if !matches!(object_type, 0 | 1 | 2 | 5) {
                    return None;
                }
                let name = if object_type == 0 {
                    None
                } else {
                    Some(directory_entry_name(entry)?)
                };
                let stream_size = read_u64(entry, 120)?;
                if self.major_version == 3 && stream_size > u64::from(u32::MAX) {
                    return None;
                }
                entries.push(DirectoryEntry {
                    object_type,
                    name,
                    left_sibling: read_u32(entry, 68)?,
                    right_sibling: read_u32(entry, 72)?,
                    child: read_u32(entry, 76)?,
                    start_sector: read_u32(entry, 116)?,
                    stream_size,
                });
            }
            directory_sector = self.fat_entry(directory_sector)?;
        }

        if directory_sector != END_OF_CHAIN
            || self
                .declared_directory_sector_count
                .is_some_and(|declared| declared != directory_sector_count)
        {
            return None;
        }
        Some(entries)
    }

    fn fat_entry(&self, sector: u32) -> Option<u32> {
        let sector = usize::try_from(sector).ok()?;
        let entries_per_fat_sector = self.sector_size.checked_div(size_of::<u32>())?;
        let fat_sector_index = sector.checked_div(entries_per_fat_sector)?;
        let entry_index = sector.checked_rem(entries_per_fat_sector)?;
        let fat_sector = *self.fat_sector_ids.get(fat_sector_index)?;
        let bytes = sector_bytes(self.input, self.sector_size, fat_sector)?;
        read_u32(bytes, entry_index * size_of::<u32>())
    }
}

fn directory_entry_name(entry: &[u8]) -> Option<String> {
    let name_length = usize::from(read_u16(entry, 64)?);
    if !(2..=64).contains(&name_length) || name_length % size_of::<u16>() != 0 {
        return None;
    }
    let name_bytes = entry.get(..name_length)?;
    if read_u16(name_bytes, name_length - size_of::<u16>())? != 0 {
        return None;
    }
    let units = name_bytes[..name_length - size_of::<u16>()]
        .chunks_exact(size_of::<u16>())
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&units).ok()
}

fn is_regular_sector(sector: u32, sector_count: usize) -> bool {
    usize::try_from(sector).is_ok_and(|sector| sector < sector_count)
}

fn sector_bytes(input: &[u8], sector_size: usize, sector: u32) -> Option<&[u8]> {
    let sector = usize::try_from(sector).ok()?;
    let start = sector.checked_add(1)?.checked_mul(sector_size)?;
    input.get(start..start.checked_add(sector_size)?)
}

fn read_u16(input: &[u8], offset: usize) -> Option<u16> {
    let bytes = input.get(offset..offset.checked_add(size_of::<u16>())?)?;
    Some(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(input: &[u8], offset: usize) -> Option<u32> {
    let bytes = input.get(offset..offset.checked_add(size_of::<u32>())?)?;
    Some(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn read_u64(input: &[u8], offset: usize) -> Option<u64> {
    let bytes = input.get(offset..offset.checked_add(size_of::<u64>())?)?;
    Some(u64::from_le_bytes([
        bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
    ]))
}
